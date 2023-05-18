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
using System.Xml.Linq;
using Aspose.Cells;
using System.IO;

namespace SHGraduationWarning.UIForm
{
    public partial class frmMain : BaseForm
    {
        string AllStr = "全部";
        // 選擇Tab名稱
        string SelectedTabName = "";

        string GWTabName = "畢業預警";
        string ChkEditTabName = "資料合理檢查_科目級別";
        string ChkEditTabName2 = "資料合理檢查_科目屬性";
        Workbook wb;
        Dictionary<string, int> _ColIdxDict;

        // 載入預設資料
        BackgroundWorker bgWorkerLoadDefault;
        // 載入畢業資格檢查資料
        BackgroundWorker bgwDataGWLoad;
        // 載入資料合理-科目級別
        BackgroundWorker bgwDataChkEditLoad;
        // 載入資料合理-科目屬性
        BackgroundWorker bgwDataChkEditLoad2;

        // 資料合理檢查報告
        BackgroundWorker bgwDataChkEditReport;

        // 處理報部科目名稱更新使用
        BackgroundWorker bwDataUpdateDSubject;

        List<ClassDeptInfo> ClassDeptInfoList;
        string SelectedGradeYearYear = "3";
        string SelectedDeptName = "";
        string SelectedClassName = "";

        Dictionary<string, List<string>> DeptNameDict;
        Dictionary<string, List<string>> GradeYearDeptNameDict;
        Dictionary<string, List<string>> GradeYearClassNameDict;

        // 檢查有問題資料 -- 科目級別
        List<StudSubjectInfo> StudSubjectInfoList;


        // 檢查有問題資料 -- 科目屬性(依課規分類)
        Dictionary<string, StudSchoolYearSubjectNameInfo> StudSchoolYearSubjectNameDict;

        // 課程規劃表對照
        Dictionary<string, GPlanInfo> GPlanDict;

        // 檢查有問題資料 -- 科目級別
        Dictionary<string, StudSubjectInfo> hasErrorSubjectInfoDict;

        // 科別名稱與編號對照
        Dictionary<string, string> DeptNameIDDic;

        // 班級名稱與編號對照
        Dictionary<string, string> ClassNameIDDic;

        // 更新科目
        List<StudSubjectInfo> UpdateSubjectInfoList;

        // 刪除科目
        List<StudSubjectInfo> DeleteSubjectInfoList;

        //        string SelectedTextName = "";

        // 資料合理檢查報告--報表資料用
        List<DataRow> chkDataReport2;
        List<DataRow> chkDataReport3;

        // 資料合理檢查報告--科目屬性使用
        List<DataRow> chkDataReport4;

        // 資料合理檢查報告--報部科目名稱使用
        List<DataRow> chkDataDSubject;

        public frmMain()
        {
            _ColIdxDict = new Dictionary<string, int>();
            GPlanDict = new Dictionary<string, GPlanInfo>();
            UpdateSubjectInfoList = new List<StudSubjectInfo>();
            DeleteSubjectInfoList = new List<StudSubjectInfo>();
            hasErrorSubjectInfoDict = new Dictionary<string, StudSubjectInfo>();
            DeptNameIDDic = new Dictionary<string, string>();
            ClassNameIDDic = new Dictionary<string, string>();
            wb = new Workbook();

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

            bgwDataChkEditLoad2 = new BackgroundWorker();
            bgwDataChkEditLoad2.DoWork += BgwDataChkEditLoad2_DoWork;
            bgwDataChkEditLoad2.RunWorkerCompleted += BgwDataChkEditLoad2_RunWorkerCompleted;
            bgwDataChkEditLoad2.ProgressChanged += BgwDataChkEditLoad2_ProgressChanged;
            bgwDataChkEditLoad2.WorkerReportsProgress = true;

            bwDataUpdateDSubject = new BackgroundWorker();
            bwDataUpdateDSubject.DoWork += BwDataUpdateDSubject_DoWork;
            bwDataUpdateDSubject.RunWorkerCompleted += BwDataUpdateDSubject_RunWorkerCompleted;
            bwDataUpdateDSubject.ProgressChanged += BwDataUpdateDSubject_ProgressChanged;
            bwDataUpdateDSubject.WorkerReportsProgress = true;

            StudSubjectInfoList = new List<StudSubjectInfo>();

            StudSchoolYearSubjectNameDict = new Dictionary<string, StudSchoolYearSubjectNameInfo>();

            DeptNameDict = new Dictionary<string, List<string>>();
            ClassDeptInfoList = new List<ClassDeptInfo>(); ;
            SelectedDeptName = SelectedClassName = AllStr;
            chkDataReport2 = new List<DataRow>();
            chkDataReport3 = new List<DataRow>();
            chkDataReport4 = new List<DataRow>();
            GradeYearDeptNameDict = new Dictionary<string, List<string>>();
            GradeYearClassNameDict = new Dictionary<string, List<string>>();

            bgwDataChkEditReport = new BackgroundWorker();
            bgwDataChkEditReport.DoWork += BgwDataChkEditReport_DoWork;
            bgwDataChkEditReport.ProgressChanged += BgwDataChkEditReport_ProgressChanged;
            bgwDataChkEditReport.RunWorkerCompleted += BgwDataChkEditReport_RunWorkerCompleted;
            bgwDataChkEditReport.WorkerReportsProgress = true;

            chkDataDSubject = new List<DataRow>();


            InitializeComponent();
        }

        private void BwDataUpdateDSubject_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("報部科目名稱更新中...", e.ProgressPercentage);
        }

        private void BwDataUpdateDSubject_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
            if (e.Error == null)
            {
                int count = (int)e.Result;
                MsgBox.Show("共更新" + count + "筆。");
                // 清除報部科目暫存
                chkDataDSubject.Clear();

            }
            else
            {
                MsgBox.Show(e.Error.Message);
            }

            ControlEnable(true);

        }

        private void BwDataUpdateDSubject_DoWork(object sender, DoWorkEventArgs e)
        {
            bwDataUpdateDSubject.ReportProgress(10);
            int UpdateCount = 0;
            if (chkDataDSubject.Count > 0)
                UpdateCount = DataAccess.UpdateSemsScoreDSubjectName(chkDataDSubject);

            e.Result = UpdateCount;

            bwDataUpdateDSubject.ReportProgress(100);
        }

        private void BgwDataChkEditReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
            if (SelectedTabName == ChkEditTabName)
            {
                if (wb != null)
                {
                    Utility.ExportXls("學期成績與課規比對", wb);
                }
            }

            if (SelectedTabName == ChkEditTabName2)
            {
                if (wb != null)
                {
                    // 因為匯入學期科目成績目前只支援2003格式
                    Utility.ExportXls2003("學期科目成績匯入檔", wb);
                }
            }

            btnReport.Enabled = true;
        }

        private void BgwDataChkEditReport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("報表產生中...", e.ProgressPercentage);
        }

        private void BgwDataChkEditReport_DoWork(object sender, DoWorkEventArgs e)
        {
            if (SelectedTabName == ChkEditTabName)
            {
                // 產生報表
                try
                {
                    bgwDataChkEditReport.ReportProgress(1);
                    // 填值到 Excel
                    wb = new Workbook(new MemoryStream(Properties.Resources.學期成績與課規比對樣板));
                    Worksheet wstSC = wb.Worksheets["依學期成績為主比對課規不符合"];
                    wstSC.Name = "依學期成績為主比對課規不符合";

                    Worksheet wstSC2 = wb.Worksheets["依課規為主比對學期成績不符合"];
                    wstSC2.Name = "依課規為主比對學期成績不符合";

                    Worksheet wstSC3 = wb.Worksheets["依課規比對課程群組學分總數不符合"];
                    wstSC3.Name = "依課規比對課程群組學分總數不符合";

                    int rowIdx = 1;
                    _ColIdxDict.Clear();
                    // 讀取欄位與索引            
                    for (int co = 0; co <= wstSC.Cells.MaxDataColumn; co++)
                    {
                        _ColIdxDict.Add(wstSC.Cells[0, co].StringValue, co);
                    }


                    if (hasErrorSubjectInfoDict.Count > 0)
                    {
                        foreach (StudSubjectInfo ss in hasErrorSubjectInfoDict.Values)
                        {
                            wstSC.Cells[rowIdx, GetColIndex("學生系統編號")].PutValue(ss.StudentID);
                            wstSC.Cells[rowIdx, GetColIndex("學號")].PutValue(ss.StudentNumber);

                            wstSC.Cells[rowIdx, GetColIndex("班級")].PutValue(ss.ClassName);
                            wstSC.Cells[rowIdx, GetColIndex("座號")].PutValue(ss.SeatNo);
                            wstSC.Cells[rowIdx, GetColIndex("科別")].PutValue(ss.DeptName);
                            wstSC.Cells[rowIdx, GetColIndex("姓名")].PutValue(ss.Name);
                            wstSC.Cells[rowIdx, GetColIndex("學年度")].PutValue(ss.SchoolYear);
                            wstSC.Cells[rowIdx, GetColIndex("學期")].PutValue(ss.Semester);
                            wstSC.Cells[rowIdx, GetColIndex("成績年級")].PutValue(ss.GradeYear);
                            wstSC.Cells[rowIdx, GetColIndex("科目名稱")].PutValue(ss.SubjectName);
                            wstSC.Cells[rowIdx, GetColIndex("科目級別")].PutValue(ss.SubjectLevel);
                            wstSC.Cells[rowIdx, GetColIndex("新科目名稱")].PutValue(ss.SubjectNameNew);
                            wstSC.Cells[rowIdx, GetColIndex("新科目級別")].PutValue(ss.SubjectLevelNew);
                            wstSC.Cells[rowIdx, GetColIndex("使用課規")].PutValue(ss.GPName);
                            wstSC.Cells[rowIdx, GetColIndex("學分數")].PutValue(ss.Credit);
                            wstSC.Cells[rowIdx, GetColIndex("不計學分")].PutValue(ss.NotIncludedInCredit);
                            wstSC.Cells[rowIdx, GetColIndex("不需評分")].PutValue(ss.NotIncludedInCalc);
                            //wstSC.Cells[rowIdx, GetColIndex("問題說明")].PutValue(string.Join(",", ss.ErrorMsgList.ToArray()));
                            rowIdx++;
                        }
                        wstSC.AutoFitColumns();
                    }

                    bgwDataChkEditReport.ReportProgress(40);

                    rowIdx = 1;
                    _ColIdxDict.Clear();
                    // 讀取欄位與索引            
                    for (int co = 0; co <= wstSC2.Cells.MaxDataColumn; co++)
                    {
                        _ColIdxDict.Add(wstSC2.Cells[0, co].StringValue, co);
                    }

                    if (chkDataReport2.Count > 0)
                    {
                        foreach (DataRow dr in chkDataReport2)
                        {
                            wstSC2.Cells[rowIdx, GetColIndex("學號")].PutValue(dr["學號"] + "");
                            wstSC2.Cells[rowIdx, GetColIndex("科別")].PutValue(dr["科別名稱"] + "");
                            wstSC2.Cells[rowIdx, GetColIndex("班級")].PutValue(dr["班級"] + "");
                            wstSC2.Cells[rowIdx, GetColIndex("座號")].PutValue(dr["座號"] + "");
                            wstSC2.Cells[rowIdx, GetColIndex("姓名")].PutValue(dr["姓名"] + "");
                            wstSC2.Cells[rowIdx, GetColIndex("使用課程規劃表")].PutValue(dr["使用課程規劃表"] + "");
                            wstSC2.Cells[rowIdx, GetColIndex("學期")].PutValue(dr["學期"] + "");
                            wstSC2.Cells[rowIdx, GetColIndex("領域")].PutValue(dr["領域"] + "");
                            wstSC2.Cells[rowIdx, GetColIndex("科目名稱")].PutValue(dr["科目名稱"] + "");
                            wstSC2.Cells[rowIdx, GetColIndex("科目級別")].PutValue(dr["科目級別"] + "");
                            rowIdx++;
                        }
                    }

                    wstSC2.AutoFitColumns();
                    bgwDataChkEditReport.ReportProgress(70);

                    rowIdx = 1;
                    _ColIdxDict.Clear();
                    // 讀取欄位與索引            
                    for (int co = 0; co <= wstSC3.Cells.MaxDataColumn; co++)
                    {
                        _ColIdxDict.Add(wstSC3.Cells[0, co].StringValue, co);
                    }

                    if (chkDataReport3.Count > 0)
                    {
                        foreach (DataRow dr in chkDataReport3)
                        {
                            wstSC3.Cells[rowIdx, GetColIndex("學號")].PutValue(dr["學號"] + "");
                            wstSC3.Cells[rowIdx, GetColIndex("科別")].PutValue(dr["科別名稱"] + "");
                            wstSC3.Cells[rowIdx, GetColIndex("班級")].PutValue(dr["班級"] + "");
                            wstSC3.Cells[rowIdx, GetColIndex("座號")].PutValue(dr["座號"] + "");
                            wstSC3.Cells[rowIdx, GetColIndex("姓名")].PutValue(dr["姓名"] + "");
                            wstSC3.Cells[rowIdx, GetColIndex("使用課程規劃表")].PutValue(dr["使用課程規劃表"] + "");
                            wstSC3.Cells[rowIdx, GetColIndex("成績年級")].PutValue(dr["成績年級"] + "");
                            wstSC3.Cells[rowIdx, GetColIndex("學期")].PutValue(dr["學期"] + "");
                            wstSC3.Cells[rowIdx, GetColIndex("分組名稱")].PutValue(dr["分組名稱"] + "");
                            wstSC3.Cells[rowIdx, GetColIndex("分組修課學分數")].PutValue(dr["分組修課學分數"] + "");
                            wstSC3.Cells[rowIdx, GetColIndex("成績累計學分數")].PutValue(dr["成績累計學分數"] + "");
                            rowIdx++;
                        }
                    }
                    wstSC3.AutoFitColumns();
                    bgwDataChkEditReport.ReportProgress(100);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            // 資料合理檢查-科目屬性
            if (SelectedTabName == ChkEditTabName2)
            {
                try
                {
                    bgwDataChkEditReport.ReportProgress(10);
                    // 填值到 Excel
                    wb = new Workbook(new MemoryStream(Properties.Resources.學期成績匯入檔樣板));
                    Worksheet wstSC = wb.Worksheets["學期科目成績匯入檔"];
                    wstSC.Name = "學期科目成績匯入檔";

                    int rowIdx = 1;
                    _ColIdxDict.Clear();
                    // 讀取欄位與索引            
                    for (int co = 0; co <= wstSC.Cells.MaxDataColumn; co++)
                    {
                        _ColIdxDict.Add(wstSC.Cells[0, co].StringValue, co);
                    }


                    if (chkDataReport4.Count > 0)
                    {
                        foreach (DataRow dr in chkDataReport4)
                        {
                            wstSC.Cells[rowIdx, GetColIndex("學生系統編號")].PutValue(dr["學生系統編號"] + "");
                            wstSC.Cells[rowIdx, GetColIndex("學號")].PutValue(dr["學號"] + "");
                            wstSC.Cells[rowIdx, GetColIndex("科別")].PutValue(dr["科別名稱"] + "");
                            wstSC.Cells[rowIdx, GetColIndex("班級")].PutValue(dr["班級"] + "");
                            wstSC.Cells[rowIdx, GetColIndex("座號")].PutValue(dr["座號"] + "");
                            wstSC.Cells[rowIdx, GetColIndex("姓名")].PutValue(dr["姓名"] + "");
                            wstSC.Cells[rowIdx, GetColIndex("學年度")].PutValue(dr["學年度"] + "");
                            wstSC.Cells[rowIdx, GetColIndex("學期")].PutValue(dr["學期"] + "");
                            wstSC.Cells[rowIdx, GetColIndex("學期")].PutValue(dr["學期"] + "");
                            wstSC.Cells[rowIdx, GetColIndex("成績年級")].PutValue(dr["成績年級"] + "");
                            wstSC.Cells[rowIdx, GetColIndex("科目")].PutValue(dr["科目名稱"] + "");
                            wstSC.Cells[rowIdx, GetColIndex("科目級別")].PutValue(dr["科目級別"] + "");
                            wstSC.Cells[rowIdx, GetColIndex("領域")].PutValue(dr["新領域"] + "");
                            wstSC.Cells[rowIdx, GetColIndex("指定學年科目名稱")].PutValue(dr["新指定學年科目名稱"] + "");
                            wstSC.Cells[rowIdx, GetColIndex("課程代碼")].PutValue(dr["新課程代碼"] + "");
                            wstSC.Cells[rowIdx, GetColIndex("分項類別")].PutValue(dr["新分項類別"] + "");
                            wstSC.Cells[rowIdx, GetColIndex("必選修")].PutValue(dr["新必選修"] + "");

                            // 轉換部訂字為部定
                            string RequiredBy = dr["新校部訂"] + "";
                            if (RequiredBy == "部訂")
                                RequiredBy = "部定";

                            wstSC.Cells[rowIdx, GetColIndex("校部訂")].PutValue(RequiredBy);
                            wstSC.Cells[rowIdx, GetColIndex("報部科目名稱")].PutValue(dr["新報部科目名稱"] + "");

                            rowIdx++;
                        }

                        wstSC.AutoFitColumns();
                    }


                    bgwDataChkEditReport.ReportProgress(100);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

        }

        private void BgwDataChkEditLoad2_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("資料合理檢查中...", e.ProgressPercentage);
        }

        private void BgwDataChkEditLoad2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
            dgData2ChkEdit.Rows.Clear();

            // 取得學期成績與課規以科目名稱+級別比對相同，分項類別、領域、校部訂、必選修、指定學年科目名稱、課程代碼、報部科目名稱，不同。
            if (chkDataReport4.Count > 0)
            {
                foreach (DataRow dr in chkDataReport4)
                {
                    int rowIdx = dgData2ChkEdit.Rows.Add();
                    dgData2ChkEdit.Rows[rowIdx].Cells["學號"].Value = dr["學號"] + "";
                    dgData2ChkEdit.Rows[rowIdx].Cells["科別"].Value = dr["科別名稱"] + "";
                    dgData2ChkEdit.Rows[rowIdx].Cells["班級"].Value = dr["班級"] + "";
                    dgData2ChkEdit.Rows[rowIdx].Cells["座號"].Value = dr["座號"] + "";
                    dgData2ChkEdit.Rows[rowIdx].Cells["姓名"].Value = dr["姓名"] + "";
                    dgData2ChkEdit.Rows[rowIdx].Cells["學年度"].Value = dr["學年度"] + "";
                    dgData2ChkEdit.Rows[rowIdx].Cells["學期"].Value = dr["學期"] + "";
                    dgData2ChkEdit.Rows[rowIdx].Cells["成績年級"].Value = dr["成績年級"] + "";
                    dgData2ChkEdit.Rows[rowIdx].Cells["科目名稱"].Value = dr["科目名稱"] + "";
                    dgData2ChkEdit.Rows[rowIdx].Cells["科目級別"].Value = dr["科目級別"] + "";
                    dgData2ChkEdit.Rows[rowIdx].Cells["領域"].Value = dr["領域"] + "";
                    dgData2ChkEdit.Rows[rowIdx].Cells["新領域"].Value = dr["新領域"] + "";
                    dgData2ChkEdit.Rows[rowIdx].Cells["指定學年科目名稱"].Value = dr["指定學年科目名稱"] + "";
                    dgData2ChkEdit.Rows[rowIdx].Cells["新指定學年科目名稱"].Value = dr["新指定學年科目名稱"] + "";
                    dgData2ChkEdit.Rows[rowIdx].Cells["課程代碼"].Value = dr["課程代碼"] + "";
                    dgData2ChkEdit.Rows[rowIdx].Cells["新課程代碼"].Value = dr["新課程代碼"] + "";
                    dgData2ChkEdit.Rows[rowIdx].Cells["分項類別"].Value = dr["分項類別"] + "";
                    dgData2ChkEdit.Rows[rowIdx].Cells["新分項類別"].Value = dr["新分項類別"] + "";
                    // 轉換部訂字為部定
                    string RequiredBy = dr["校部訂"] + "";
                    if (RequiredBy == "部訂")
                        RequiredBy = "部定";

                    dgData2ChkEdit.Rows[rowIdx].Cells["校部訂"].Value = RequiredBy;

                    string RequiredByNew = dr["新校部訂"] + "";
                    if (RequiredByNew == "部訂")
                        RequiredByNew = "部定";

                    dgData2ChkEdit.Rows[rowIdx].Cells["新校部訂"].Value = RequiredByNew;

                    dgData2ChkEdit.Rows[rowIdx].Cells["必選修"].Value = dr["必選修"] + "";
                    dgData2ChkEdit.Rows[rowIdx].Cells["新必選修"].Value = dr["新必選修"] + "";
                    dgData2ChkEdit.Rows[rowIdx].Cells["報部科目名稱"].Value = dr["報部科目名稱"] + "";
                }
            }

            lblMsg.Text = "共" + dgData2ChkEdit.Rows.Count + "筆";
            ControlEnable(true);
            btnDel.Enabled = false;
        }

        private void BgwDataChkEditLoad2_DoWork(object sender, DoWorkEventArgs e)
        {
            int rpInt = 10;
            bgwDataChkEditLoad2.ReportProgress(rpInt);

            string DeptID = "";
            string ClassID = "";

            if (DeptNameIDDic.ContainsKey(SelectedDeptName))
                DeptID = DeptNameIDDic[SelectedDeptName];

            if (ClassNameIDDic.ContainsKey(SelectedClassName))
                ClassID = ClassNameIDDic[SelectedClassName];

            // 取得學期成績與課規以科目名稱+級別比對相同，領域、指定學年科目名稱、課程代碼、分項、校部定、必選修、報部科目，不同。
            chkDataReport4 = DataAccess.GetSemsSubjectLevelCheckGraduationPlan4(SelectedGradeYearYear, DeptID, ClassID);


            rpInt = 70;
            bgwDataChkEditLoad2.ReportProgress(rpInt);
            // 取得報部科目名稱，更新使用。
            chkDataDSubject = DataAccess.GetSemsSubjectLevelCheckGraduationPlan5(SelectedGradeYearYear, DeptID, ClassID);

            rpInt = 100;
            bgwDataChkEditLoad2.ReportProgress(rpInt);

        }

        private void BgwDataChkEditLoad_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("資料合理檢查中...", e.ProgressPercentage);
        }

        private void BgwDataChkEditLoad_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");

            try
            {
                dgDataChkEdit.Rows.Clear();
                if (hasErrorSubjectInfoDict.Count > 0)
                {
                    foreach (StudSubjectInfo ss in hasErrorSubjectInfoDict.Values)
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
                        dgDataChkEdit.Rows[rowIdx].Cells["新科目名稱"].Value = ss.SubjectNameNew;
                        dgDataChkEdit.Rows[rowIdx].Cells["新科目級別"].Value = ss.SubjectLevelNew;
                        //dgDataChkEdit.Rows[rowIdx].Cells["分項"].Value = ss.Entry;
                        //dgDataChkEdit.Rows[rowIdx].Cells["領域"].Value = ss.Domain;
                        //dgDataChkEdit.Rows[rowIdx].Cells["學分"].Value = ss.Credit;
                        //dgDataChkEdit.Rows[rowIdx].Cells["必選修"].Value = ss.Required;
                        //dgDataChkEdit.Rows[rowIdx].Cells["校部定"].Value = ss.RequiredBy;
                        dgDataChkEdit.Rows[rowIdx].Cells["使用課規"].Value = ss.GPName;
                        //dgDataChkEdit.Rows[rowIdx].Cells["指定學年科目名稱"].Value = ss.SchoolYearSubjectName;
                        dgDataChkEdit.Rows[rowIdx].Cells["問題說明"].Value = string.Join(",", ss.ErrorMsgList.ToArray());
                        //if (ss.IsSubjectLevelChanged && ss.SubjectLevelNew != "")
                        //    dgDataChkEdit.Rows[rowIdx].Cells["勾選"].Value = "是";
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


            lblMsg.Text = "共" + dgDataChkEdit.Rows.Count + "筆";
            ControlEnable(true);
        }

        private void BgwDataChkEditLoad_DoWork(object sender, DoWorkEventArgs e)
        {
            int rpInt = 1;
            StudSubjectInfoList.Clear();
            GPlanDict.Clear();
            hasErrorSubjectInfoDict.Clear();
            bgwDataChkEditLoad.ReportProgress(rpInt);
            chkDataReport2.Clear();
            chkDataReport3.Clear();

            string DeptID = "";
            string ClassID = "";

            if (DeptNameIDDic.ContainsKey(SelectedDeptName))
                DeptID = DeptNameIDDic[SelectedDeptName];

            if (ClassNameIDDic.ContainsKey(SelectedClassName))
                ClassID = ClassNameIDDic[SelectedClassName];


            // 檢查學期成績與課規比對
            StudSubjectInfoList.AddRange(DataAccess.GetSemsSubjectLevelCheckGraduationPlan1(SelectedGradeYearYear, DeptID, ClassID));
            rpInt = 30;
            bgwDataChkEditLoad.ReportProgress(rpInt);

            // 檢查課規與學期成績比對
            chkDataReport2 = DataAccess.GetSemsSubjectLevelCheckGraduationPlan2(SelectedGradeYearYear, DeptID, ClassID);


            rpInt = 70;
            bgwDataChkEditLoad.ReportProgress(rpInt);

            // 檢查課規課程群組學分總數不符合
            chkDataReport3 = DataAccess.GetSemsSubjectLevelCheckGraduationPlan3(SelectedGradeYearYear, DeptID, ClassID);

            // 填入科目級別不同
            foreach (StudSubjectInfo ss in StudSubjectInfoList)
            {
                ss.ErrorMsgList.Add("科目級別與課規不同");
                AddErrorSubjectInfoDict(ss);
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
            int rpInt = 1;
            bgwDataGWLoad.ReportProgress(rpInt);



            rpInt = 100;
            bgwDataGWLoad.ReportProgress(rpInt);
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
            btnReport.Enabled = btnDel.Enabled = btnUpdate.Enabled = false;
        }

        private void BgWorkerLoadDefault_DoWork(object sender, DoWorkEventArgs e)
        {
            int rpInt = 1;
            bgWorkerLoadDefault.ReportProgress(rpInt);

            // 年級科別名稱對照
            GradeYearDeptNameDict.Clear();
            GradeYearClassNameDict.Clear();
            DeptNameIDDic.Clear();
            ClassNameIDDic.Clear();

            // 讀取班級科別資訊            
            ClassDeptInfoList = DataAccess.GetClassDeptList();
            DeptNameDict.Clear();
            foreach (ClassDeptInfo cd in ClassDeptInfoList)
            {
                if (!DeptNameDict.ContainsKey(cd.DeptName))
                    DeptNameDict.Add(cd.DeptName, new List<string>());

                DeptNameDict[cd.DeptName].Add(cd.ClassName);

                if (!ClassNameIDDic.ContainsKey(cd.ClassName))
                    ClassNameIDDic.Add(cd.ClassName, cd.ClassID);

                if (!DeptNameIDDic.ContainsKey(cd.DeptName))
                    DeptNameIDDic.Add(cd.DeptName, cd.DeptID);

                if (!GradeYearDeptNameDict.ContainsKey(cd.GradeYear))
                    GradeYearDeptNameDict.Add(cd.GradeYear, new List<string>());

                if (!GradeYearClassNameDict.ContainsKey(cd.GradeYear))
                    GradeYearClassNameDict.Add(cd.GradeYear, new List<string>());

                GradeYearDeptNameDict[cd.GradeYear].Add(cd.DeptName);
                GradeYearClassNameDict[cd.GradeYear].Add(cd.ClassName);
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
            if (!comboDept.Items.Contains(AllStr))
                comboDept.Items.Remove(AllStr);

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
            // 更新報部科目名稱按鈕
            buttonUpdateDSubjectName.Visible = false;
            ClearClassDept();
            SelectedTabName = ChkEditTabName;
            LoadTabDesc();

            comboGradeYear.Items.Add("3");
            comboGradeYear.Items.Add("2");
            comboGradeYear.Items.Add("1");
            comboGradeYear.Text = "3";

            comboGradeYear.DropDownStyle = ComboBoxStyle.DropDownList;
            comboDept.DropDownStyle = ComboBoxStyle.DropDownList;
            comboClass.DropDownStyle = ComboBoxStyle.DropDownList;

            // 載入畢業預警欄位
            //    LoadDgGWColumns();

            // 載入資料合理性檢查欄位--科目級別
            LoadDgDataChkColumns();

            // 載入資料合理性檢查欄位--科目屬性
            LoadDgData2ChkColumns();

            tabControl1.SelectedTabIndex = 0;

            // 讀取班級科別資訊
            bgWorkerLoadDefault.RunWorkerAsync();

        }

        private void ControlEnable(bool value)
        {
            btnQuery.Enabled = comboDept.Enabled = comboClass.Enabled = btnReport.Enabled = value;

            buttonUpdateDSubjectName.Enabled = value;

            tabControl1.Enabled = value;
            btnDel.Enabled = btnUpdate.Enabled = false;
        }

        private void comboDept_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 清除班級，填入相對值
            comboClass.Text = "";
            SelectedClassName = "";
            SelectedDeptName = comboDept.Text;

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
                    comboClass.Items.Add(AllStr);
                    foreach (string name in DeptNameDict[SelectedDeptName])
                    {
                        if (GradeYearClassNameDict.ContainsKey(SelectedGradeYearYear))
                        {
                            if (GradeYearClassNameDict[SelectedGradeYearYear].Contains(name))
                                comboClass.Items.Add(name);
                        }

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

        // 載入資料合理檢查欄位--科目級別
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
                    ""Width"": 30,
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
                    ""HeaderText"": ""新科目名稱"",
                    ""Name"": ""新科目名稱"",
                    ""Width"": 120,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""新科目級別"",
                    ""Name"": ""新科目級別"",
                    ""Width"": 30,
                    ""ReadOnly"": true
                },                
                {
                    ""HeaderText"": ""使用課規"",
                    ""Name"": ""使用課規"",
                    ""Width"": 120,
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


                //// 加入刪除勾選
                //DataGridViewCheckBoxColumn chkCol1 = new DataGridViewCheckBoxColumn();
                //chkCol1.Name = "勾選";
                //chkCol1.HeaderText = "勾選";
                //chkCol1.Width = 30;
                //chkCol1.TrueValue = "是";
                //chkCol1.FalseValue = "否";
                //chkCol1.IndeterminateValue = "否";
                //dgDataChkEdit.Columns.Add(chkCol1);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        // 載入資料合理檢查欄位--科目屬性
        private void LoadDgData2ChkColumns()
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
                    ""Width"": 30,
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
                    ""HeaderText"": ""分項類別"",
                    ""Name"": ""分項類別"",
                    ""Width"": 80,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""新分項類別"",
                    ""Name"": ""新分項類別"",
                    ""Width"": 80,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""領域"",
                    ""Name"": ""領域"",
                    ""Width"": 80,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""新領域"",
                    ""Name"": ""新領域"",
                    ""Width"": 80,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""校部訂"",
                    ""Name"": ""校部訂"",
                    ""Width"": 40,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""新校部訂"",
                    ""Name"": ""新校部訂"",
                    ""Width"": 40,
                    ""ReadOnly"": true
                }
                ,
                {
                    ""HeaderText"": ""必選修"",
                    ""Name"": ""必選修"",
                    ""Width"": 40,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""新必選修"",
                    ""Name"": ""新必選修"",
                    ""Width"": 40,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""指定學年科目名稱"",
                    ""Name"": ""指定學年科目名稱"",
                    ""Width"": 120,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""新指定學年科目名稱"",
                    ""Name"": ""新指定學年科目名稱"",
                    ""Width"": 120,
                    ""ReadOnly"": true
                },                
                {
                    ""HeaderText"": ""課程代碼"",
                    ""Name"": ""課程代碼"",
                    ""Width"": 120,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""新課程代碼"",
                    ""Name"": ""新課程代碼"",
                    ""Width"": 120,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""報部科目名稱"",
                    ""Name"": ""報部科目名稱"",
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
                    dgData2ChkEdit.Columns.Add(dgt);
                }


                //// 加入刪除勾選
                //DataGridViewCheckBoxColumn chkCol1 = new DataGridViewCheckBoxColumn();
                //chkCol1.Name = "勾選";
                //chkCol1.HeaderText = "勾選";
                //chkCol1.Width = 30;
                //chkCol1.TrueValue = "是";
                //chkCol1.FalseValue = "否";
                //chkCol1.IndeterminateValue = "否";

                //dgData2ChkEdit.Columns.Add(chkCol1);

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
            // 資料合理檢查 -- 科目級別
            if (SelectedTabName == ChkEditTabName)
            {
                DeleteSubjectInfoList.Clear();
                foreach (DataGridViewRow drv in dgDataChkEdit.Rows)
                {
                    if (drv.Cells["勾選"].Value != null)
                        if (drv.Cells["勾選"].Value.ToString() == "是")
                        {
                            // MsgBox.Show("刪除");
                            StudSubjectInfo ssi = drv.Tag as StudSubjectInfo;
                            if (ssi != null)
                                DeleteSubjectInfoList.Add(ssi);
                        }
                }
                if (DeleteSubjectInfoList.Count > 0)
                {

                    if (FISCA.Presentation.Controls.MsgBox.Show("選「是」將刪除勾選" + DeleteSubjectInfoList.Count + "筆學期科目資料，請確認? ", "刪除學期科目", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                    {
                        int delCount = DataAccess.DelSemsScoreSubject(DeleteSubjectInfoList);
                        MsgBox.Show("已刪除" + delCount + "筆資料。");

                        dgDataChkEdit.Rows.Clear();
                        lblMsg.Text = "共0筆";
                        btnUpdate.Enabled = btnDel.Enabled = false;
                    }
                }
            }
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            // 資料合理檢查 -- 科目級別
            if (SelectedTabName == ChkEditTabName)
            {
                UpdateSubjectInfoList.Clear();
                foreach (DataGridViewRow drv in dgDataChkEdit.Rows)
                {
                    StudSubjectInfo ssi = drv.Tag as StudSubjectInfo;
                    if (ssi != null)
                    {
                        if (ssi.IsSubjectLevelChanged)
                        {
                            if (drv.Cells["勾選"].Value != null)
                                if (drv.Cells["勾選"].Value.ToString() == "是")
                                {
                                    UpdateSubjectInfoList.Add(ssi);
                                }
                        }

                    }
                }
                if (UpdateSubjectInfoList.Count > 0)
                {
                    if (FISCA.Presentation.Controls.MsgBox.Show("選「是」將更新勾選" + UpdateSubjectInfoList.Count + "筆學期科目資料，請確認? ", "更新學期科目", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                    {
                        int UpdateCount = DataAccess.UpdateSemsScoreSubjectInfo(UpdateSubjectInfoList);
                        MsgBox.Show("已更新" + UpdateCount + "筆資料。");

                        dgDataChkEdit.Rows.Clear();
                        lblMsg.Text = "共0筆";
                        btnUpdate.Enabled = btnDel.Enabled = false;
                    }
                }
            }
        }
        private void btnQuery_Click(object sender, EventArgs e)
        {
            // 畢業預警
            if (SelectedTabName == GWTabName)
            {
                ControlEnable(false);
                bgwDataGWLoad.RunWorkerAsync();
            }

            // 資料合理檢查
            if (SelectedTabName == ChkEditTabName)
            {
                ControlEnable(false);
                bgwDataChkEditLoad.RunWorkerAsync();
            }

            // 資料合理檢查-科目屬性
            if (SelectedTabName == ChkEditTabName2)
            {
                ControlEnable(false);
                bgwDataChkEditLoad2.RunWorkerAsync();
            }

        }

        private void tabControl1_SelectedTabChanged(object sender, DevComponents.DotNetBar.TabStripTabChangedEventArgs e)
        {

        }

        private void AddErrorSubjectInfoDict(StudSubjectInfo ssi)
        {
            string key = ssi.SchoolYear + "_" + ssi.Semester + "_" + ssi.StudentID + "_" + ssi.SubjectName + "_" + ssi.SubjectLevel;
            if (!hasErrorSubjectInfoDict.ContainsKey(key))
                hasErrorSubjectInfoDict.Add(key, ssi);
            else
            {
                if (ssi.ErrorMsgList.Count > 0)
                {
                    foreach (string str in ssi.ErrorMsgList)
                    {
                        if (!hasErrorSubjectInfoDict[key].ErrorMsgList.Contains(str))
                            hasErrorSubjectInfoDict[key].ErrorMsgList.Add(str);
                    }
                }
            }
        }


        private void comboClass_SelectedIndexChanged(object sender, EventArgs e)
        {
            SelectedClassName = comboClass.Text;
        }

        private void chkItemAll_CheckedChanged(object sender, EventArgs e)
        {
            // 資料合理檢查 -- 科目級別
            if (SelectedTabName == ChkEditTabName)
            {
                foreach (DataGridViewRow drv in dgDataChkEdit.Rows)
                {
                    //drv.Cells["勾選"].Value = chkItemAll.Checked ? "是" : "否";
                }
            }
        }

        private void tbItemGW_Click(object sender, EventArgs e)
        {
            // 畢業預警
            SelectedTabName = GWTabName;
            btnQuery.Enabled = false;
            buttonUpdateDSubjectName.Visible = false;
            LoadTabDesc();
            lblMsg.Text = "共0筆";
        }

        private void tbItemChkEdit_Click(object sender, EventArgs e)
        {
            // 資料合理檢查 -- 科目級別
            SelectedTabName = ChkEditTabName;
            buttonUpdateDSubjectName.Visible = false;
            btnQuery.Enabled = true;
            LoadTabDesc();
            lblMsg.Text = "共" + dgDataChkEdit.Rows.Count + "筆";
        }



        // 載入說明資訊
        private string LoadTabDesc()
        {
            string msg = "";
            // 畢業預警
            if (SelectedTabName == GWTabName)
            {
                msg = @"
";
            }


            else if (SelectedTabName == ChkEditTabName)
            {
                // 資料合理檢查_科目級別
                msg = @"學生成績資料科目級別合理性檢查：
1.檢查範圍：一般及延修狀態學生科目。
2.依年級、科別、班級條件，檢查範圍內學生學期成績進行合理性檢查。以學生學期科目成績之科目名稱+級別比對學生採用之課程規劃表產生3張工作表：
    1.依學期成績為主比對課規不符合：列出與課規比對不符合的清單，並針對科目名稱比對的到的資料提供新的科目名稱和新的級別。
    2.依課規為主比對學期成績不符合：列出與課規比對後學生學期科目缺少的資料清單。
    3.依課規比對課程群組學分總數不符合：列出與課規中設定的課程群組之學分數不足之資料清單。
";

            }
            else if (SelectedTabName == ChkEditTabName2)
            {
                // 資料合理檢查_科目屬性
                msg = @"資料合理檢查(科目屬性)：(學生範圍：學生狀態：一般+延修)
1. 檢查學期成績與課程規劃，以科目名稱+科目級別比對相同，找出分項類別、領域、校部訂、必選修、指定學年科目名稱、課程代碼、報部科目名稱 有差異資料。
2. 產生報表:產生可匯入學期成績檔案。
3. 更新報部科目名稱：依學生課程規劃比對學期成績，當科目名稱+級別相同，將學期成績的報部科目名稱更新為課程規劃報部科目名稱。
";
            }

            return msg;
        }

        private void tabControl1_Click(object sender, EventArgs e)
        {

        }

        private void tbItem2ChkEdit_Click(object sender, EventArgs e)
        {
            // 資料合理檢查 -- 科目屬性
            SelectedTabName = ChkEditTabName2;
            btnQuery.Enabled = true;
            buttonUpdateDSubjectName.Visible = true;
            LoadTabDesc();
            //lblMsg.Text = "共" + dgData2ChkEdit.Rows.Count + "筆";
        }

        private void btnReport_Click(object sender, EventArgs e)
        {
            if (SelectedTabName == ChkEditTabName)
            {
                btnReport.Enabled = false;
                bgwDataChkEditReport.RunWorkerAsync();
            }

            if (SelectedTabName == ChkEditTabName2)
            {
                btnReport.Enabled = false;
                bgwDataChkEditReport.RunWorkerAsync();
            }
        }

        private int GetColIndex(string name)
        {
            int value = 0;
            if (_ColIdxDict.ContainsKey(name))
                value = _ColIdxDict[name];

            return value;
        }

        private void comboGradeYear_SelectedIndexChanged(object sender, EventArgs e)
        {

            // 清除科別班級，填入相對值

            comboDept.Text = "";
            SelectedDeptName = "";
            comboClass.Text = "";
            SelectedClassName = "";

            SelectedGradeYearYear = comboGradeYear.Text;

            comboClass.Items.Clear();
            comboDept.Items.Clear();

            if (!comboDept.Items.Contains(AllStr))
                comboDept.Items.Add(AllStr);

            if (GradeYearDeptNameDict.ContainsKey(SelectedGradeYearYear))
            {
                foreach (string name in GradeYearDeptNameDict[SelectedGradeYearYear])
                {
                    if (!comboDept.Items.Contains(name))
                        comboDept.Items.Add(name);
                }
            }
            comboDept.Text = AllStr;
            SelectedDeptName = AllStr;

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
                    comboClass.Items.Add(AllStr);
                    foreach (string name in DeptNameDict[SelectedDeptName])
                    {
                        if (GradeYearClassNameDict.ContainsKey(SelectedGradeYearYear))
                        {
                            if (GradeYearClassNameDict[SelectedGradeYearYear].Contains(name))
                                comboClass.Items.Add(name);
                        }

                    }
                    if (comboClass.Items.Count > 0)
                        comboClass.SelectedIndex = 0;
                }
            }
        }

        private void buttonUpdateDSubjectName_Click(object sender, EventArgs e)
        {
            buttonUpdateDSubjectName.Enabled = false;
            if (chkDataDSubject.Count == 0)
            {
                MsgBox.Show("沒有資料可更新");
                buttonUpdateDSubjectName.Enabled = true;
            }
            else
            {
                if (MsgBox.Show("選「是」將更新" + chkDataDSubject.Count + "筆報部科目名稱。", "更新報部科目名稱", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    ControlEnable(false);
                    // 更新報部科目名稱
                    bwDataUpdateDSubject.RunWorkerAsync();
                }
            }
        }

        private void buttonLoadDesc_Click(object sender, EventArgs e)
        {
            string desc = LoadTabDesc();
            frmDesc frm = new frmDesc();
            frm.SetDesc(desc);
            frm.ShowDialog();
        }
    }
}
