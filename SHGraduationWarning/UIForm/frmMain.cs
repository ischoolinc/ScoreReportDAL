﻿using System;
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
using SmartSchool.Customization.Data;
using K12.Data;
using System.Xml;
using SmartSchool.Evaluation.Reports.MultiSemesterScore.DataModel;
using Aspose.Words;
using System.Security.Cryptography;

namespace SHGraduationWarning.UIForm
{
    public partial class frmMain : BaseForm
    {
        string AllStr = "全部";
        string NoGradeYearStr = "未分年級";
        // 選擇Tab名稱
        string SelectedTabName = "";

        // 打勾符號
        string chkMark = "✔";

        string GWTabName = "畢業預警";
        string ChkEditTabName = "資料合理檢查_科目級別";
        string ChkEditTabName2 = "資料合理檢查_科目屬性";
        Workbook wb;
        Dictionary<string, int> _ColIdxDict;

        // 樣板設定
        Configure configure;

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

        // 學生報表合併使用
        DataTable StudDT;

        // 處理畢業預警列印報表
        BackgroundWorker bgwGrandCheckReport;

        // 處理畢業預警列印報表(班級)
        BackgroundWorker bgwGrandCheckClassReport;

        // 處理畢業預警學生清單匯出
        BackgroundWorker bgwGrandCheckStudentExport;


        List<ClassDeptInfo> ClassDeptInfoList;
        string SelectedGradeYearYear = "3";
        string SelectedDeptName = "";
        string SelectedClassName = "";

        // 畢業預警學生報表資料(StudentID,ReportStudentInfo)
        Dictionary<string, ReportStudentInfo> ReportStudentDict;
        List<ReportStudentInfo> ReportStudentList;

        List<ReportStudentInfo> SelectedReportStudentList;

        // 畢業預警班級報表資料(ReportClassInfo)
        Dictionary<string, ReportClassInfo> ReportClassDict;
        List<ReportClassInfo> ReportClassList;

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

        // 僅顯示未達畢業標準
        bool isChkNotUptoGStandard = false;
        public frmMain()
        {

            _ColIdxDict = new Dictionary<string, int>();
            GPlanDict = new Dictionary<string, GPlanInfo>();
            UpdateSubjectInfoList = new List<StudSubjectInfo>();
            DeleteSubjectInfoList = new List<StudSubjectInfo>();
            hasErrorSubjectInfoDict = new Dictionary<string, StudSubjectInfo>();
            DeptNameIDDic = new Dictionary<string, string>();
            ClassNameIDDic = new Dictionary<string, string>();
            ReportStudentDict = new Dictionary<string, ReportStudentInfo>();
            ReportStudentList = new List<ReportStudentInfo>();
            ReportClassList = new List<ReportClassInfo>();
            ReportClassDict = new Dictionary<string, ReportClassInfo>();
            SelectedReportStudentList = new List<ReportStudentInfo>();

            wb = new Workbook();
            StudDT = new DataTable();

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

            bgwGrandCheckReport = new BackgroundWorker();
            bgwGrandCheckReport.DoWork += BgwGrandCheckReport_DoWork;
            bgwGrandCheckReport.RunWorkerCompleted += BgwGrandCheckReport_RunWorkerCompleted;
            bgwGrandCheckReport.ProgressChanged += BgwGrandCheckReport_ProgressChanged;
            bgwGrandCheckReport.WorkerReportsProgress = true;

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

            bgwGrandCheckClassReport = new BackgroundWorker();
            bgwGrandCheckClassReport.DoWork += BgwGrandCheckClassReport_DoWork;
            bgwGrandCheckClassReport.RunWorkerCompleted += BgwGrandCheckClassReport_RunWorkerCompleted;
            bgwGrandCheckClassReport.ProgressChanged += BgwGrandCheckClassReport_ProgressChanged;
            bgwGrandCheckClassReport.WorkerReportsProgress = true;

            bgwGrandCheckStudentExport = new BackgroundWorker();
            bgwGrandCheckStudentExport.DoWork += BgwGrandCheckStudentExport_DoWork;
            bgwGrandCheckStudentExport.RunWorkerCompleted += BgwGrandCheckStudentExport_RunWorkerCompleted;
            bgwGrandCheckStudentExport.ProgressChanged += BgwGrandCheckStudentExport_ProgressChanged;
            bgwGrandCheckStudentExport.WorkerReportsProgress = true;

            InitializeComponent();
        }

        private void BgwGrandCheckStudentExport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("畢業預警學生清單匯出中...", e.ProgressPercentage);

        }

        private void BgwGrandCheckStudentExport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");

            if (e.Error == null)
            {
                try
                {
                    Workbook wb1 = e.Result as Workbook;
                    if (wb1 != null)
                    {
                        Utility.ExportXls("畢業預警學生清單", wb1);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            else
            {
                MsgBox.Show(e.Error.Message);
            }

            ControlEnable(true);
        }

        private void BgwGrandCheckStudentExport_DoWork(object sender, DoWorkEventArgs e)
        {
            bgwGrandCheckStudentExport.ReportProgress(20);
            // 處理報表填入 Excel
            Workbook wb = new Workbook(new MemoryStream(Properties.Resources.畢業審查學生清單樣板));
            Worksheet wst = wb.Worksheets["學生清單"];

            int rowIdx = 1;
            _ColIdxDict.Clear();
            // 讀取欄位與索引            
            for (int co = 0; co <= wst.Cells.MaxDataColumn; co++)
            {
                _ColIdxDict.Add(wst.Cells[0, co].StringValue, co);
            }

            foreach (ReportStudentInfo rs in ReportStudentList)
            {
                // 畢業預警審查通過不顯示
                if (isChkNotUptoGStandard && rs.GraduationCheck == "通過")
                    continue;

                // 學號	班級	座號	姓名	畢業審查
                wst.Cells[rowIdx, _ColIdxDict["學號"]].PutValue(rs.StudentNumber);
                wst.Cells[rowIdx, _ColIdxDict["班級"]].PutValue(rs.ClassName);
                wst.Cells[rowIdx, _ColIdxDict["座號"]].PutValue(rs.SeatNo);
                wst.Cells[rowIdx, _ColIdxDict["姓名"]].PutValue(rs.StudentName);
                if (rs.GraduationCheck == "通過")
                {
                    wst.Cells[rowIdx, _ColIdxDict["畢業審查"]].PutValue("符合畢業標準");
                }
                else if (rs.GraduationCheck == "不通過")
                {
                    wst.Cells[rowIdx, _ColIdxDict["畢業審查"]].PutValue("未達畢業標準");
                }
                else
                {
                    wst.Cells[rowIdx, _ColIdxDict["畢業審查"]].PutValue("");
                }
                rowIdx++;
            }
            wst.AutoFitColumns();
            e.Result = wb;

            bgwGrandCheckStudentExport.ReportProgress(100);
        }

        private void BgwGrandCheckClassReport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("班級畢業預警報表產生中...", e.ProgressPercentage);
        }

        private void BgwGrandCheckClassReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");

            if (e.Error == null)
            {
                try
                {
                    Workbook wb1 = e.Result as Workbook;
                    if (wb1 != null)
                    {
                        Utility.ExportXls("班級預警報表", wb1);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            else
            {
                MsgBox.Show(e.Error.Message);
            }

            ControlEnable(true);
        }

        private void BgwGrandCheckClassReport_DoWork(object sender, DoWorkEventArgs e)
        {

            // 清空班級暫存資料
            foreach (string id in ReportClassDict.Keys)
            {
                ReportClassDict[id].dicClassRule.Clear();
                ReportClassDict[id].StudentInfoList.Clear();
            }

            // 產生班級畢業預警報表
            bgwGrandCheckClassReport.ReportProgress(5);

            // 檢查班級與學生是否有資料，並判斷僅產生未達畢業
            if (ReportClassList.Count > 0 && ReportStudentList.Count > 0)
            {
                foreach (ReportStudentInfo rs in ReportStudentList)
                {
                    // 畢業預警審查通過不顯示
                    if (isChkNotUptoGStandard && rs.GraduationCheck == "通過")
                        continue;

                    // 整理班級資料
                    if (ReportClassDict.ContainsKey(rs.ClassID))
                    {
                        ReportClassDict[rs.ClassID].StudentInfoList.Add(rs);
                        // 處理班級共通規則
                        foreach (XmlElement xmlRule in rs.GraGrandCheckXml.SelectNodes("畢業規則"))
                        {
                            if (xmlRule.GetAttribute("啟用") == "是")
                            {
                                string key = "";
                                string Rule = xmlRule.GetAttribute("規則");
                                if (Rule == "學年學業成績及格" || Rule == "功過相抵未滿三大過")
                                {
                                    key = Rule;
                                }

                                if (xmlRule.GetAttribute("核心科目表序號") == "")
                                {
                                    // 班級共通規則
                                    if (Rule == "應修總學分數")
                                        key = "總學分";
                                    if (Rule == "總學分數")
                                        key = "總學分";
                                    if (Rule == "應修所有必修課程")
                                        key = "必修";
                                    if (Rule == "必修學分數")
                                        key = "必修";
                                    if (Rule == "應修所有部定必修課程")
                                        key = "部定必修";
                                    if (Rule == "部訂必修學分數")
                                        key = "部定必修";
                                    if (Rule == "校訂必修學分數")
                                        key = "校定必修";

                                    if (Rule == "選修學分數")
                                        key = "選修";

                                    if (Rule == "應修專業及實習總學分數")
                                        key = "專業及實習";

                                    if (Rule == "專業及實習總學分數")
                                        key = "專業及實習";

                                    if (Rule == "實習學分數")
                                        key = "實習";

                                }
                                else
                                {
                                    // 自訂
                                    key = xmlRule.GetAttribute("核心科目表名稱");

                                }
                                if (!ReportClassDict[rs.ClassID].dicClassRule.ContainsKey(key))
                                    ReportClassDict[rs.ClassID].dicClassRule.Add(key, new List<string>());
                                if (!ReportClassDict[rs.ClassID].dicClassRule[key].Contains(Rule))
                                    ReportClassDict[rs.ClassID].dicClassRule[key].Add(Rule);
                            }
                        }

                        // 將規則內應修排在前面，主要避免先讀到取得再讀到應修，產生錯誤。
                        foreach (string key in ReportClassDict[rs.ClassID].dicClassRule.Keys)
                        {
                            ReportClassDict[rs.ClassID].dicClassRule[key].Sort((x, y) =>
                            {
                                if (x.Contains("應修") && !y.Contains("應修"))
                                    return -1;
                                else if (!x.Contains("應修") && y.Contains("應修"))
                                    return 1;
                                else
                                    return string.Compare(x, y);
                            });
                        }
                    }

                }

                bgwGrandCheckClassReport.ReportProgress(50);


                // 處理報表填入 Excel
                Workbook wb = new Workbook(new MemoryStream(Properties.Resources.班級預警報表樣板));
                // 用來複製用樣板
                // temp
                Worksheet wstTemp = wb.Worksheets["temp"];

                // 定義固定排列順序
                List<string> itemSorter = new List<string> { "總學分", "必修", "部定必修", "校定必修", "選修", "專業及實習", "實習" };

                //  新增樣板
                foreach (string cid in ReportClassDict.Keys)
                {
                    // 沒有學生不處理
                    if (ReportClassDict[cid].StudentInfoList.Count == 0)
                        continue;

                    string ClassName = ReportClassDict[cid].ClassName;
                    // 新增Excel 工作表(班級名稱)
                    Worksheet wstNew = wb.Worksheets.Add(ClassName);

                    // 姓名
                    wstNew.Cells.CopyColumns(wstTemp.Cells, 0, 0, 3);

                    // 處理樣板項目固定順訊總學分先
                    List<string> ItemNameList = new List<string>();
                    foreach (string name in ReportClassDict[cid].dicClassRule.Keys)
                    {
                        // 放最後不參與排序
                        if (name == "學年學業成績及格" || name == "功過相抵未滿三大過")
                            continue;

                        if (itemSorter.Contains(name))
                            ItemNameList.Add(name);
                    }

                    // 將固定排序
                    ItemNameList.Sort((x, y) =>
                        itemSorter.IndexOf(x).CompareTo(itemSorter.IndexOf(y)));

                    // 加入自訂
                    foreach (string name in ReportClassDict[cid].dicClassRule.Keys)
                    {
                        // 放最後不參與排序
                        if (name == "學年學業成績及格" || name == "功過相抵未滿三大過")
                            continue;

                        if (!itemSorter.Contains(name))
                            ItemNameList.Add(name);
                    }


                    int ColIdx = 3;
                    // 動態樣板
                    foreach (string item in ItemNameList)
                    {
                        if (ReportClassDict[cid].dicClassRule.ContainsKey(item))
                        {
                            if (ReportClassDict[cid].dicClassRule[item].Count == 1)
                            {
                                wstNew.Cells.CopyColumns(wstTemp.Cells, 7, ColIdx, 2);
                                ColIdx += 2;
                            }
                            if (ReportClassDict[cid].dicClassRule[item].Count == 2)
                            {
                                wstNew.Cells.CopyColumns(wstTemp.Cells, 3, ColIdx, 4);
                                ColIdx += 4;
                            }
                        }
                    }


                    if (ReportClassDict[cid].dicClassRule.ContainsKey("學年學業成績及格"))
                    {
                        // 學年
                        wstNew.Cells.CopyColumns(wstTemp.Cells, 9, ColIdx, 1);
                        ColIdx++;
                    }

                    if (ReportClassDict[cid].dicClassRule.ContainsKey("功過相抵未滿三大過"))
                    {
                        // 功過
                        wstNew.Cells.CopyColumns(wstTemp.Cells, 10, ColIdx, 1);
                    }


                    // 最大值
                    ReportClassDict[cid].MaxColumnIndex = ColIdx;

                    // 學生資料
                    int studRIdx = 5;
                    foreach (ReportStudentInfo rs in ReportClassDict[cid].StudentInfoList)
                    {
                        // 複製格式
                        wstNew.Cells.CopyRows(wstNew.Cells, 5, studRIdx, 1);
                        studRIdx++;
                    }

                    // 學校
                    wstNew.Cells.CopyRows(wstTemp.Cells, 0, 0, 1);
                    // 學校合併
                    wstNew.Cells.Merge(0, 0, 1, ReportClassDict[cid].MaxColumnIndex + 1);

                    // 班級
                    wstNew.Cells.CopyRows(wstTemp.Cells, 1, 1, 1);
                    wstNew.Cells.Merge(1, 0, 1, ReportClassDict[cid].MaxColumnIndex + 1);

                    // 學校名稱
                    wstNew.Cells[0, 0].PutValue(K12.Data.School.ChineseName);
                    wstNew.Cells[1, 0].PutValue(ClassName + " 畢業預警表");

                    // 設定字靠右對齊
                    Aspose.Cells.Style ss = wstNew.Cells[2, 0].GetStyle();
                    ss.HorizontalAlignment = TextAlignmentType.Right;
                    wstNew.Cells[2, ReportClassDict[cid].MaxColumnIndex].SetStyle(ss);

                    wstNew.Cells[2, ReportClassDict[cid].MaxColumnIndex].PutValue("班導師：" + ReportClassDict[cid].TeacherName);

                    int col1 = 3;
                    foreach (string item in ItemNameList)
                    {
                        if (ReportClassDict[cid].dicClassRule.ContainsKey(item))
                        {
                            if (ReportClassDict[cid].dicClassRule[item].Count == 1)
                            {
                                string rule = ReportClassDict[cid].dicClassRule[item][0];
                                if (rule.Contains("應修"))
                                {
                                    wstNew.Cells[4, col1].PutValue("應修");
                                    wstNew.Cells[4, col1 + 1].PutValue("已修");
                                }
                                else
                                {
                                    wstNew.Cells[4, col1].PutValue("應取得");
                                    wstNew.Cells[4, col1 + 1].PutValue("取得");
                                }

                                wstNew.Cells[3, col1].PutValue(item); col1 += 2;
                            }
                            if (ReportClassDict[cid].dicClassRule[item].Count == 2)
                            {
                                foreach (string rule in ReportClassDict[cid].dicClassRule[item])
                                {
                                    wstNew.Cells[4, col1].PutValue("應修");
                                    wstNew.Cells[4, col1 + 1].PutValue("已修");
                                    wstNew.Cells[4, col1 + 2].PutValue("應取得");
                                    wstNew.Cells[4, col1 + 3].PutValue("取得");
                                }

                                wstNew.Cells[3, col1].PutValue(item); col1 += 4;

                            }

                        }
                    }

                    studRIdx = 5;
                    int rptColIdx = 0;

                    foreach (ReportStudentInfo rs in ReportClassDict[cid].StudentInfoList)
                    {
                        wstNew.Cells[studRIdx, rptColIdx].PutValue(rs.StudentNumber); rptColIdx++;
                        wstNew.Cells[studRIdx, rptColIdx].PutValue(rs.SeatNo); rptColIdx++;
                        wstNew.Cells[studRIdx, rptColIdx].PutValue(rs.StudentName); rptColIdx++;

                        foreach (string item in ItemNameList)
                        {
                            if (ReportClassDict[cid].dicClassRule.ContainsKey(item))
                            {
                                foreach (string rule in ReportClassDict[cid].dicClassRule[item])
                                {
                                    XmlNode node = rs.GraGrandCheckXml.SelectSingleNode("//畢業規則[@規則='" + rule + "']");

                                    if (node != null)
                                    {
                                        double x1, x2;
                                        if (double.TryParse(node.Attributes["通過標準"].Value, out x1))
                                        {
                                            wstNew.Cells[studRIdx, rptColIdx].PutValue(x1);
                                        }

                                        if (double.TryParse(node.Attributes["累計學分"].Value, out x2))
                                        {
                                            wstNew.Cells[studRIdx, rptColIdx + 1].PutValue(x2);
                                        }
                                    }

                                    rptColIdx += 2;
                                }
                            }
                        }

                        //foreach (string name in ReportClassDict[cid].dicClassRule.Keys)
                        //{


                        //}

                        // 資料整理
                        foreach (XmlElement xmlRule in rs.GraGrandCheckXml.SelectNodes("畢業規則"))
                        {
                            if (xmlRule.GetAttribute("啟用") == "是")
                            {
                                string Rule = xmlRule.GetAttribute("規則");


                                // 處理 學年學業成績及格
                                if (Rule == "學年學業成績及格")
                                {

                                    if (xmlRule.GetAttribute("畢業審查") == "通過")
                                        wstNew.Cells[studRIdx, ReportClassDict[cid].MaxColumnIndex - 1].PutValue("是");
                                    else if (xmlRule.GetAttribute("畢業審查") == "不通過")
                                    {
                                        wstNew.Cells[studRIdx, ReportClassDict[cid].MaxColumnIndex - 1].PutValue("否");
                                    }
                                    else
                                        wstNew.Cells[studRIdx, ReportClassDict[cid].MaxColumnIndex - 1].PutValue("");

                                }


                                // 處理 功過相抵未滿三大過
                                if (Rule == "功過相抵未滿三大過")
                                {

                                    if (xmlRule.GetAttribute("畢業審查") == "通過")
                                        wstNew.Cells[studRIdx, ReportClassDict[cid].MaxColumnIndex].PutValue("是");
                                    else if (xmlRule.GetAttribute("畢業審查") == "不通過")
                                    {
                                        wstNew.Cells[studRIdx, ReportClassDict[cid].MaxColumnIndex].PutValue("否");
                                    }
                                    else
                                        wstNew.Cells[studRIdx, ReportClassDict[cid].MaxColumnIndex].PutValue("");

                                }

                            }
                        }
                        studRIdx++;
                        rptColIdx = 0;
                    }



                }

                // 移除複製用樣板
                wb.Worksheets.RemoveAt("temp");
                e.Result = wb;

                //                Console.WriteLine("");
            }
            bgwGrandCheckClassReport.ReportProgress(100);
        }

        private void BgwGrandCheckReport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("畢業預警報表產生中...", e.ProgressPercentage);
        }

        private void BgwGrandCheckReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");

            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message);
                return;
            }

            try
            {
                Document doc = e.Result as Document;

                string reportName = K12.Data.School.DefaultSchoolYear + "學年度第" + K12.Data.School.DefaultSemester + "學生畢業預警通知單";

                string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);
                path = Path.Combine(path, reportName + ".docx");

                if (File.Exists(path))
                {
                    int i = 1;
                    while (true)
                    {
                        string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                        if (!File.Exists(newPath))
                        {
                            path = newPath;
                            break;
                        }
                    }
                }

                try
                {
                    Document document = doc;
                    document.Save(path, Aspose.Words.SaveFormat.Docx);
                    System.Diagnostics.Process.Start(path);
                }
                catch
                {
                    System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                    sd.Title = "另存新檔";
                    sd.FileName = reportName + ".docx";
                    sd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
                    if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        try
                        {
                            Document document = doc;
                            document.Save(sd.FileName, Aspose.Words.SaveFormat.Docx);

                        }
                        catch
                        {
                            FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
                doc = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
                MsgBox.Show(ex.Message);
            }

            this.btnReport.Enabled = true;

        }

        private void BgwGrandCheckReport_DoWork(object sender, DoWorkEventArgs e)
        {
            bgwGrandCheckReport.ReportProgress(10);
            StudDT.Clear();
            StudDT.Columns.Clear();
            // 建立 Word 合併欄位表
            StudDT.Columns.Clear();
            StudDT.Columns.Add("目前學年度");
            StudDT.Columns.Add("目前學期");
            StudDT.Columns.Add("學校名稱");
            StudDT.Columns.Add("科別");
            StudDT.Columns.Add("班級");
            StudDT.Columns.Add("座號");
            StudDT.Columns.Add("姓名");
            StudDT.Columns.Add("學號");
            StudDT.Columns.Add("成績計算規則");
            StudDT.Columns.Add("課程規劃表");
            StudDT.Columns.Add("畢業審查");

            // 規則
            List<string> ruList = new List<string>();
            ruList.Add("應修總學分數");
            ruList.Add("應修所有必修課程");
            ruList.Add("應修所有部定必修課程");
            ruList.Add("應修專業及實習總學分數");
            ruList.Add("總學分數");
            ruList.Add("必修學分數");
            ruList.Add("部訂必修學分數");
            ruList.Add("校訂必修學分數");
            ruList.Add("選修學分數");
            ruList.Add("專業及實習總學分數");
            ruList.Add("實習學分數");

            List<string> ruColList = new List<string>();
            List<string> ruColList1 = new List<string>();
            ruColList.Add("設定值");
            ruColList.Add("課規總學分數");
            ruColList.Add("通過標準");
            ruColList.Add("累計學分");
            ruColList1.Add("畢業差額");
            ruColList1.Add("不足學分");
            ruColList1.Add("已修習");
            ruColList1.Add("已取得");
            ruColList1.Add("尚未開課");
            ruColList1.Add("可重修");
            ruColList1.Add("可補修");
            ruColList1.Add("未修習");

            foreach (string r1 in ruList)
            {
                foreach (string r2 in ruColList)
                {
                    StudDT.Columns.Add(r1 + "_" + r2);
                }
                foreach (string r2 in ruColList1)
                {
                    StudDT.Columns.Add(r1 + "_" + r2);
                }
            }

            // 產生核心科目合併欄位，分2類：
            // 修課學分數統計
            for (int i = 1; i <= 5; i++)
            {
                StudDT.Columns.Add("修課學分數統計_核心科目表序號" + i + "_名稱");
                StudDT.Columns.Add("修課學分數統計_核心科目表序號" + i + "_規則");
                foreach (string r2 in ruColList)
                {
                    StudDT.Columns.Add("修課學分數統計_核心科目表序號" + i + "_" + r2);
                }
                foreach (string r2 in ruColList1)
                {
                    StudDT.Columns.Add("修課學分數統計_核心科目表序號" + i + "_" + r2);
                }
            }

            // 取得學分數統計
            for (int i = 1; i <= 5; i++)
            {
                StudDT.Columns.Add("取得學分數統計_核心科目表序號" + i + "_名稱");
                StudDT.Columns.Add("取得學分數統計_核心科目表序號" + i + "_規則");
                foreach (string r2 in ruColList)
                {
                    StudDT.Columns.Add("取得學分數統計_核心科目表序號" + i + "_" + r2);
                }
                foreach (string r2 in ruColList1)
                {
                    StudDT.Columns.Add("取得學分數統計_核心科目表序號" + i + "_" + r2);
                }
            }

            // 功過相抵未滿三大過
            List<string> ru3List = new List<string>();
            ru3List.Add("設定值");
            ru3List.Add("通過標準");
            ru3List.Add("目前累計支數");
            ru3List.Add("畢業審查");
            foreach (string r3 in ru3List)
                StudDT.Columns.Add("功過相抵未滿三大過_" + r3);

            // 學年學業成績及格
            List<string> ru4List = new List<string>();
            ru4List.Add("設定值");
            ru4List.Add("通過標準");
            ru4List.Add("不及格學年");
            ru4List.Add("畢業審查");


            // 處理科目
            List<string> colN1List = new List<string>();
            colN1List.Add("狀態");
            colN1List.Add("修課學年度");
            colN1List.Add("修課學期");
            colN1List.Add("科目名稱");
            colN1List.Add("科目級別");
            colN1List.Add("學分數");
            int maxSubj = 50;
            for (int s = 1; s <= maxSubj; s++)
            {
                foreach (string cname in colN1List)
                {
                    StudDT.Columns.Add("科目" + s + "_" + cname);

                }

            }

            // 處理固定規則檢查，可補考、可重修 符合打勾欄位
            // // 科目1_應修總學分數_可補考重修_打勾
            foreach (string name in ruList)
            {
                for (int i = 1; i <= maxSubj; i++)
                {
                    StudDT.Columns.Add("科目" + i + "_" + name + "_可補修重修_打勾");
                }
            }

            // 核心科目表科目符合規則可補修可重修打勾
            for (int i = 1; i <= 5; i++)
            {
                for (int j = 1; j <= maxSubj; j++)
                {
                    StudDT.Columns.Add("科目" + j + "_修課學分數統計_核心科目表序號" + i + "_規則_可補修重修_打勾");
                }
            }
            for (int i = 1; i <= 5; i++)
            {
                for (int j = 1; j <= 30; j++)
                {
                    StudDT.Columns.Add("科目" + j + "_取得學分數統計_核心科目表序號" + i + "_規則_可補修重修_打勾");
                }
            }


            // 填資料,將原本畫面上所有學生換成所選學生 SelectedReportStudentList
            if (SelectedReportStudentList.Count > 0)
            {
                // 填資料至 DataTable
                foreach (ReportStudentInfo rs in SelectedReportStudentList)
                {
                    // 判斷僅顯示不通過
                    if (isChkNotUptoGStandard && rs.GraduationCheck == "通過")
                        continue;

                    DataRow dr = StudDT.NewRow();
                    dr["目前學年度"] = K12.Data.School.DefaultSchoolYear;
                    dr["目前學期"] = K12.Data.School.DefaultSemester;
                    dr["學校名稱"] = K12.Data.School.ChineseName;
                    dr["科別"] = rs.GraGrandCheckXml.GetAttribute("科別");
                    dr["班級"] = rs.GraGrandCheckXml.GetAttribute("班級");
                    dr["座號"] = rs.GraGrandCheckXml.GetAttribute("座號");
                    dr["學號"] = rs.GraGrandCheckXml.GetAttribute("學號");
                    dr["姓名"] = rs.GraGrandCheckXml.GetAttribute("姓名");
                    dr["課程規劃表"] = rs.GraGrandCheckXml.GetAttribute("課程規劃表");
                    dr["成績計算規則"] = rs.GraGrandCheckXml.GetAttribute("成績計算規則");
                    if (rs.GraGrandCheckXml.GetAttribute("畢業審查") == "通過")
                    {
                        dr["畢業審查"] = "符合畢業標準";
                    }
                    else if (rs.GraGrandCheckXml.GetAttribute("畢業審查") == "不通過")
                    {
                        dr["畢業審查"] = "未達畢業標準";
                    }
                    else
                        dr["畢業審查"] = "";


                    //Console.WriteLine(rs.GraGrandCheckXml.OuterXml);
                    foreach (XmlElement xmlRule in rs.GraGrandCheckXml.SelectNodes("畢業規則"))
                    {
                        if (xmlRule.GetAttribute("啟用") == "是")
                        {
                            string Rule = xmlRule.GetAttribute("規則");
                            //處理一般
                            if (xmlRule.GetAttribute("核心科目表序號") == "")
                            {
                                // 處理固定規則統計                                
                                foreach (string ruCol in ruColList)
                                {
                                    string key = Rule + "_" + ruCol;
                                    if (StudDT.Columns.Contains(key))
                                    {
                                        dr[key] = xmlRule.GetAttribute(ruCol);
                                    }
                                }

                                // 處理 預警統計
                                foreach (XmlElement xmlRuleC in xmlRule.SelectNodes("預警統計"))
                                {
                                    foreach (string ruCol in ruColList1)
                                    {
                                        string key = Rule + "_" + ruCol;
                                        if (StudDT.Columns.Contains(key))
                                        {
                                            dr[key] = xmlRuleC.GetAttribute(ruCol);


                                            // 計算不足學分
                                            if (key.Contains("不足學分"))
                                            {
                                                if (Rule.Length > 2)
                                                {
                                                    int val1 = 0, val2 = 0;

                                                    string strVal1 = xmlRule.GetAttribute("設定值");
                                                    string strVal2 = xmlRuleC.GetAttribute("已取得");
                                                    if (Rule.Substring(0, 2) == "應修")
                                                    {
                                                        strVal2 = xmlRuleC.GetAttribute("已修習");
                                                    }

                                                    int.TryParse(strVal1, out val1);
                                                    int.TryParse(strVal2, out val2);

                                                    dr[key] = val1 - val2;
                                                }
                                            }

                                        }
                                    }
                                }

                            }
                            else
                            {
                                // 核心科目表
                                string c_key1 = xmlRule.GetAttribute("類型") + "_核心科目表序號" + xmlRule.GetAttribute("核心科目表序號") + "_名稱";
                                if (StudDT.Columns.Contains(c_key1))
                                {
                                    dr[c_key1] = xmlRule.GetAttribute("核心科目表名稱");
                                }

                                string c_key2 = xmlRule.GetAttribute("類型") + "_核心科目表序號" + xmlRule.GetAttribute("核心科目表序號") + "_規則";

                                // 建立核心科目規則對照
                                if (!rs.dicCoreSubjectRule.ContainsKey(Rule))
                                    rs.dicCoreSubjectRule.Add(Rule, c_key2);

                                if (StudDT.Columns.Contains(c_key2))
                                {
                                    dr[c_key2] = Rule;
                                }

                                // 處理固定規則統計                                
                                foreach (string ruCol in ruColList)
                                {
                                    string key = xmlRule.GetAttribute("類型") + "_核心科目表序號" + xmlRule.GetAttribute("核心科目表序號") + "_" + ruCol;
                                    if (StudDT.Columns.Contains(key))
                                    {
                                        dr[key] = xmlRule.GetAttribute(ruCol);
                                    }
                                }


                                // 處理 預警統計
                                foreach (XmlElement xmlRuleC in xmlRule.SelectNodes("預警統計"))
                                {
                                    foreach (string ruCol in ruColList1)
                                    {
                                        string key = xmlRule.GetAttribute("類型") + "_核心科目表序號" + xmlRule.GetAttribute("核心科目表序號") + "_" + ruCol;
                                        if (StudDT.Columns.Contains(key))
                                        {
                                            dr[key] = xmlRuleC.GetAttribute(ruCol);


                                            // 計算不足學分
                                            if (key.Contains("不足學分"))
                                            {
                                                if (Rule.Length > 5)
                                                {
                                                    int val1 = 0, val2 = 0;

                                                    string strVal1 = xmlRule.GetAttribute("設定值");
                                                    string strVal2 = xmlRuleC.GetAttribute("已取得");
                                                    if (Rule.Substring(0, 2) == "修課")
                                                    {
                                                        strVal2 = xmlRuleC.GetAttribute("已修習");
                                                    }

                                                    int.TryParse(strVal1, out val1);
                                                    int.TryParse(strVal2, out val2);

                                                    dr[key] = val1 - val2;
                                                }
                                            }
                                        }
                                    }
                                }

                            }



                            // 處理 功過相抵未滿三大過
                            if (Rule == "功過相抵未滿三大過")
                            {
                                foreach (string ruCol in ru3List)
                                {
                                    string key = Rule + "_" + ruCol;
                                    if (StudDT.Columns.Contains(key))
                                    {
                                        if (ruCol == "畢業審查")
                                        {
                                            if (xmlRule.GetAttribute(ruCol) == "通過")
                                                dr[key] = "是";
                                            else
                                                dr[key] = "否";
                                        }
                                        else
                                        {
                                            dr[key] = xmlRule.GetAttribute(ruCol);
                                        }
                                    }
                                }
                            }

                            // 處理 學年學業成績及格
                            if (Rule == "學年學業成績及格")
                            {
                                foreach (string ruCol in ru4List)
                                {
                                    string key = Rule + "_" + ruCol;
                                    if (StudDT.Columns.Contains(key))
                                    {
                                        if (ruCol == "畢業審查")
                                        {
                                            if (xmlRule.GetAttribute(ruCol) == "通過")
                                                dr[key] = "是";
                                            else
                                                dr[key] = "否";
                                        }
                                        else
                                        {
                                            dr[key] = xmlRule.GetAttribute(ruCol);
                                        }
                                    }
                                }
                            }

                            // 處理科目可補修、可重修
                            foreach (XmlElement xmlRuleS in xmlRule.SelectNodes("科目"))
                            {
                                if (xmlRuleS.GetAttribute("狀態") == "可補修" || xmlRuleS.GetAttribute("狀態") == "可重修")
                                {
                                    // 當不計學分跳過，判斷：學分數=0
                                    if (xmlRuleS.GetAttribute("學分數") == "0")
                                        continue;

                                    // key = 科目名稱+級別
                                    string sKey = xmlRuleS.GetAttribute("科目名稱") + "_" + xmlRuleS.GetAttribute("科目級別");
                                    if (!rs.dicRetake.ContainsKey(sKey))
                                        rs.dicRetake.Add(sKey, xmlRuleS);

                                    // 整理符合規則的科目與級別
                                    if (!rs.dicRetaleRelate.ContainsKey(Rule))
                                        rs.dicRetaleRelate.Add(Rule, new List<string>());
                                    rs.dicRetaleRelate[Rule].Add(sKey);
                                }
                            }
                        }
                    }

                    // 處理科目排序，依學年度 遞減 學期遞減
                    List<XmlElement> S = rs.dicRetake.Values.ToList();
                    List<string> SubjSort = (from s in S orderby s.GetAttribute("修課學年度") descending, s.GetAttribute("修課學期") descending select s.GetAttribute("科目名稱") + "_" + s.GetAttribute("科目級別")).ToList();


                    // 處理科目填入
                    int sKeyIdx = 1;
                    //foreach (string sKey in rs.dicRetake.Keys)
                    foreach (string sKey in SubjSort)
                    {
                        if (rs.dicRetake.ContainsKey(sKey))
                        {
                            // 處理科目屬性填入
                            foreach (string name in colN1List)
                            {
                                string sK1 = "科目" + sKeyIdx + "_" + name;
                                if (StudDT.Columns.Contains(sK1))
                                {
                                    dr[sK1] = rs.dicRetake[sKey].GetAttribute(name);
                                }
                            }

                            // 處理符合規則打勾
                            // 科目1_應修總學分數_可補修重修_打勾
                            foreach (string key in rs.dicRetaleRelate.Keys)
                            {
                                if (rs.dicRetaleRelate[key].Contains(sKey))
                                {
                                    string rKey = "科目" + sKeyIdx + "_" + key + "_可補修重修_打勾";

                                    if (StudDT.Columns.Contains(rKey))
                                    {
                                        dr[rKey] = chkMark;
                                    }

                                    // 處理核心科目規則
                                    if (rs.dicCoreSubjectRule.ContainsKey(key))
                                    {
                                        rKey = "科目" + sKeyIdx + "_" + rs.dicCoreSubjectRule[key] + "_可補修重修_打勾";
                                        if (StudDT.Columns.Contains(rKey))
                                        {
                                            dr[rKey] = chkMark;
                                        }
                                    }

                                }
                            }

                            sKeyIdx++;
                        }
                    }

                    StudDT.Rows.Add(dr);
                }


                // 合併
                if (this.configure != null)
                {
                    bgwGrandCheckReport.ReportProgress(50);
                    Document doc = configure.Template.Clone();
                    doc.MailMerge.Execute(StudDT);
                    doc.MailMerge.DeleteFields();
                    e.Result = doc;
                }
            }

            bgwGrandCheckReport.ReportProgress(100);
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
                            wstSC2.Cells[rowIdx, GetColIndex("分項類別")].PutValue(dr["新分項類別"] + "");
                            // 校部訂字轉換
                            string RequiredBy = dr["新校部訂"] + "";
                            if (RequiredBy == "部訂")
                                RequiredBy = "部定";

                            wstSC2.Cells[rowIdx, GetColIndex("校部訂")].PutValue(RequiredBy);
                            wstSC2.Cells[rowIdx, GetColIndex("必選修")].PutValue(dr["新必選修"] + "");
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
                    bgwDataChkEditReport.ReportProgress(20);

                    if (chkDataReport4.Count < 65535)
                    {
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
                    }
                    else
                    {
                        // 更換新樣板
                        // 填值到 Excel
                        wb = new Workbook(new MemoryStream(Properties.Resources.學期成績匯入檔樣板1));
                        Worksheet wstSC = wb.Worksheets["學期科目成績匯入檔"];
                        wstSC.Name = "學期科目成績匯入檔";

                        Worksheet wstSC1 = wb.Worksheets["學期科目成績匯入檔1"];
                        wstSC1.Name = "學期科目成績匯入檔1";


                        if (chkDataReport4.Count > 0)
                        {
                            int rowIdx = 1;
                            _ColIdxDict.Clear();
                            // 讀取欄位與索引            
                            for (int co = 0; co <= wstSC.Cells.MaxDataColumn; co++)
                            {
                                _ColIdxDict.Add(wstSC.Cells[0, co].StringValue, co);
                            }

                            // 放第一工作表
                            for (int drIdx = 0; drIdx < 65535; drIdx++)
                            {
                                DataRow dr = chkDataReport4[drIdx];

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

                            rowIdx = 1;
                            _ColIdxDict.Clear();
                            // 讀取欄位與索引            
                            for (int co = 0; co <= wstSC1.Cells.MaxDataColumn; co++)
                            {
                                _ColIdxDict.Add(wstSC1.Cells[0, co].StringValue, co);
                            }

                            // 放第2工作表
                            for (int drIdx = 65535; drIdx < chkDataReport4.Count; drIdx++)
                            {
                                DataRow dr = chkDataReport4[drIdx];

                                wstSC1.Cells[rowIdx, GetColIndex("學生系統編號")].PutValue(dr["學生系統編號"] + "");
                                wstSC1.Cells[rowIdx, GetColIndex("學號")].PutValue(dr["學號"] + "");
                                wstSC1.Cells[rowIdx, GetColIndex("科別")].PutValue(dr["科別名稱"] + "");
                                wstSC1.Cells[rowIdx, GetColIndex("班級")].PutValue(dr["班級"] + "");
                                wstSC1.Cells[rowIdx, GetColIndex("座號")].PutValue(dr["座號"] + "");
                                wstSC1.Cells[rowIdx, GetColIndex("姓名")].PutValue(dr["姓名"] + "");
                                wstSC1.Cells[rowIdx, GetColIndex("學年度")].PutValue(dr["學年度"] + "");
                                wstSC1.Cells[rowIdx, GetColIndex("學期")].PutValue(dr["學期"] + "");
                                wstSC1.Cells[rowIdx, GetColIndex("學期")].PutValue(dr["學期"] + "");
                                wstSC1.Cells[rowIdx, GetColIndex("成績年級")].PutValue(dr["成績年級"] + "");
                                wstSC1.Cells[rowIdx, GetColIndex("科目")].PutValue(dr["科目名稱"] + "");
                                wstSC1.Cells[rowIdx, GetColIndex("科目級別")].PutValue(dr["科目級別"] + "");
                                wstSC1.Cells[rowIdx, GetColIndex("領域")].PutValue(dr["新領域"] + "");
                                wstSC1.Cells[rowIdx, GetColIndex("指定學年科目名稱")].PutValue(dr["新指定學年科目名稱"] + "");
                                wstSC1.Cells[rowIdx, GetColIndex("課程代碼")].PutValue(dr["新課程代碼"] + "");
                                wstSC1.Cells[rowIdx, GetColIndex("分項類別")].PutValue(dr["新分項類別"] + "");
                                wstSC1.Cells[rowIdx, GetColIndex("必選修")].PutValue(dr["新必選修"] + "");

                                // 轉換部訂字為部定
                                string RequiredBy = dr["新校部訂"] + "";
                                if (RequiredBy == "部訂")
                                    RequiredBy = "部定";

                                wstSC1.Cells[rowIdx, GetColIndex("校部訂")].PutValue(RequiredBy);
                                wstSC1.Cells[rowIdx, GetColIndex("報部科目名稱")].PutValue(dr["新報部科目名稱"] + "");

                                rowIdx++;
                            }

                            wstSC1.AutoFitColumns();
                        }


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
                    int seatNo;
                    if (int.TryParse(dr["座號"] + "", out seatNo))
                    {
                        dgData2ChkEdit.Rows[rowIdx].Cells["座號"].Value = seatNo;
                    }
                    else
                    {
                        dgData2ChkEdit.Rows[rowIdx].Cells["座號"].Value = 0;
                    }

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
                    dgData2ChkEdit.Rows[rowIdx].Cells["新報部科目名稱"].Value = dr["新報部科目名稱"] + "";
                }
            }

            lblMsg.Text = "共" + dgData2ChkEdit.Rows.Count + "筆";
            ControlEnable(true);

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

            // 處理未分年級
            if (SelectedGradeYearYear == NoGradeYearStr)
            {
                chkDataReport4 = DataAccess.GetSemsSubjectLevelCheckGraduationPlan4NoGradeYear(SelectedGradeYearYear, "", ClassID);

                rpInt = 80;
                bgwDataChkEditLoad2.ReportProgress(rpInt);

                chkDataDSubject = DataAccess.GetSemsSubjectLevelCheckGraduationPlan5NoGradeYear(SelectedGradeYearYear, "", ClassID);

            }
            else
            {
                if (string.IsNullOrEmpty(DeptID))
                {
                    List<string> DeptIDList = new List<string>();
                    foreach (ClassDeptInfo ci in ClassDeptInfoList)
                    {
                        if (ci.GradeYear == SelectedGradeYearYear)
                            if (!DeptIDList.Contains(ci.DeptID))
                                DeptIDList.Add(ci.DeptID);
                    }

                    // 一個年級分科
                    chkDataReport4.Clear();
                    foreach (string id in DeptIDList)
                    {
                        chkDataReport4.AddRange(DataAccess.GetSemsSubjectLevelCheckGraduationPlan4(SelectedGradeYearYear, id, ClassID));
                    }
                }
                else
                {
                    // 單科
                    chkDataReport4 = DataAccess.GetSemsSubjectLevelCheckGraduationPlan4(SelectedGradeYearYear, DeptID, ClassID);
                }


                rpInt = 80;
                bgwDataChkEditLoad2.ReportProgress(rpInt);


            }

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
                        int seatNo;
                        if (int.TryParse(ss.SeatNo, out seatNo))
                        {
                            dgDataChkEdit.Rows[rowIdx].Cells["座號"].Value = seatNo;
                        }
                        else
                            dgDataChkEdit.Rows[rowIdx].Cells["座號"].Value = 0;


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


            // 未分年級
            if (SelectedGradeYearYear == NoGradeYearStr)
            {
                // 一般處理
                // 檢查學期成績與課規比對
                StudSubjectInfoList.AddRange(DataAccess.GetSemsSubjectLevelCheckGraduationPlan1NoGradeYear(SelectedGradeYearYear, "", ClassID));
                rpInt = 30;
                bgwDataChkEditLoad.ReportProgress(rpInt);

                // 檢查課規與學期成績比對
                chkDataReport2 = DataAccess.GetSemsSubjectLevelCheckGraduationPlan2NoGradeYear(SelectedGradeYearYear, "", ClassID);


                rpInt = 70;
                bgwDataChkEditLoad.ReportProgress(rpInt);

                // 檢查課規課程群組學分總數不符合
                chkDataReport3 = DataAccess.GetSemsSubjectLevelCheckGraduationPlan3NoGradeYearNoGradeYear(SelectedGradeYearYear, "", ClassID);

                // 填入科目級別不同
                foreach (StudSubjectInfo ss in StudSubjectInfoList)
                {
                    ss.ErrorMsgList.Add("科目級別與課規不同");
                    AddErrorSubjectInfoDict(ss);
                }
            }
            else
            {
                // 一般處理
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
            }

            rpInt = 100;
            bgwDataChkEditLoad.ReportProgress(rpInt);

        }

        private void BgwDataGWLoad_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("畢業預警檢查中...", e.ProgressPercentage);
        }

        private void BgwDataGWLoad_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
            ReloadDataGridViewGW();
            ControlEnable(true);

        }

        private void BgwDataGWLoad_DoWork(object sender, DoWorkEventArgs e)
        {
            int rpInt = 10;
            bgwDataGWLoad.ReportProgress(rpInt);
            ReportStudentDict.Clear();
            ReportStudentList.Clear();
            ReportClassDict.Clear();
            ReportClassList.Clear();
            string DeptID = "";
            string ClassID = "";

            if (DeptNameIDDic.ContainsKey(SelectedDeptName))
                DeptID = DeptNameIDDic[SelectedDeptName];

            if (ClassNameIDDic.ContainsKey(SelectedClassName))
                ClassID = ClassNameIDDic[SelectedClassName];

            // 延修生，未分年級
            if (SelectedGradeYearYear == NoGradeYearStr)
            {
                // 只有單一班級處理
                // 單科 學生
                ReportStudentList = DataAccess.GetReportStudentListN(SelectedGradeYearYear, DeptID, ClassID);

                // 單科 班級
                ReportClassList = DataAccess.GetReportClassListN(SelectedGradeYearYear, DeptID, ClassID);
            }
            else
            {
                if (string.IsNullOrEmpty(DeptID))
                {
                    List<string> DeptIDList = new List<string>();
                    foreach (ClassDeptInfo ci in ClassDeptInfoList)
                    {
                        if (ci.GradeYear == SelectedGradeYearYear)
                            if (!DeptIDList.Contains(ci.DeptID))
                                DeptIDList.Add(ci.DeptID);
                    }

                    // 一個年級分科            
                    foreach (string id in DeptIDList)
                    {
                        // 學生
                        ReportStudentList.AddRange(DataAccess.GetReportStudentList(SelectedGradeYearYear, id, ClassID));

                        // 班級
                        ReportClassList.AddRange(DataAccess.GetReportClassList(SelectedGradeYearYear, id, ClassID));
                    }
                }
                else
                {
                    // 單科 學生
                    ReportStudentList = DataAccess.GetReportStudentList(SelectedGradeYearYear, DeptID, ClassID);

                    // 單科 班級
                    ReportClassList = DataAccess.GetReportClassList(SelectedGradeYearYear, DeptID, ClassID);
                }

            }


            rpInt = 40;
            bgwDataGWLoad.ReportProgress(rpInt);

            // 整理資料
            foreach (ReportStudentInfo rs in ReportStudentList)
            {
                if (!ReportStudentDict.ContainsKey(rs.StudentID))
                    ReportStudentDict.Add(rs.StudentID, rs);
            }
            foreach (ReportClassInfo rc in ReportClassList)
            {
                if (!ReportClassDict.ContainsKey(rc.ClassID))
                    ReportClassDict.Add(rc.ClassID, rc);
            }

            // 取得學生畢業報告
            AccessHelper accessHelper = new AccessHelper();
            List<SmartSchool.Customization.Data.StudentRecord> studentRecList = accessHelper.StudentHelper.GetStudents(ReportStudentDict.Keys);
            new SmartSchool.Evaluation.WearyDogComputer().FillStudentGradCheck(accessHelper, studentRecList);
            foreach (SmartSchool.Customization.Data.StudentRecord studRec in
                studentRecList)
            {
                if (studRec.Fields.ContainsKey("GrandCheckReport"))
                {
                    XmlElement xmlElement = studRec.Fields["GrandCheckReport"] as XmlElement;
                    if (xmlElement != null)
                    {
                        if (ReportStudentDict.ContainsKey(studRec.StudentID))
                        {
                            ReportStudentDict[studRec.StudentID].GraGrandCheckXml = xmlElement;
                            // 取得畢業審查結果
                            ReportStudentDict[studRec.StudentID].GraduationCheck = xmlElement.GetAttribute("畢業審查");
                        }
                    }
                }
            }


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
            btnReport.Enabled = false;
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

                if (cd.DeptName != "")
                {
                    if (!DeptNameIDDic.ContainsKey(cd.DeptName))
                        DeptNameIDDic.Add(cd.DeptName, cd.DeptID);
                }

                if (!GradeYearDeptNameDict.ContainsKey(cd.GradeYear))
                    GradeYearDeptNameDict.Add(cd.GradeYear, new List<string>());

                if (!GradeYearClassNameDict.ContainsKey(cd.GradeYear))
                    GradeYearClassNameDict.Add(cd.GradeYear, new List<string>());

                if (!GradeYearDeptNameDict[cd.GradeYear].Contains(cd.DeptName))
                    GradeYearDeptNameDict[cd.GradeYear].Add(cd.DeptName);

                if (!GradeYearClassNameDict[cd.GradeYear].Contains(cd.ClassName))
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

            // 加入年級科別
            if (GradeYearDeptNameDict.ContainsKey(SelectedGradeYearYear))
            {
                foreach (string name in GradeYearDeptNameDict[SelectedGradeYearYear])
                    comboDept.Items.Add(name);
            }

            // 加入年級班級
            if (GradeYearClassNameDict.ContainsKey(SelectedGradeYearYear))
            {
                foreach (string name in GradeYearClassNameDict[SelectedGradeYearYear])
                {
                    comboClass.Items.Add(name);
                }

            }

            // 判斷是否有未分年級再加入
            if (GradeYearClassNameDict.ContainsKey(NoGradeYearStr))
            {
                if (!comboGradeYear.Items.Contains(NoGradeYearStr))
                    comboGradeYear.Items.Add(NoGradeYearStr);
            }

            comboDept.Text = AllStr;
            comboClass.Text = AllStr;
            this.ResumeLayout();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {

            ControlEnable(false);
            lnkSetReportTemplate.Visible = false;
            ChkNotUptoGStandard.Visible = false;
            btnExport.Visible = btnClassReport.Visible = false;

            ClearClassDept();
            SelectedTabName = ChkEditTabName;
            LoadTabDesc();

            // 取得目前一般狀態學生的年級
            try
            {
                List<string> grList = DataAccess.GetStudentGradeYear12();
                if (grList.Count > 0)
                {
                    comboGradeYear.Items.Clear();
                    foreach (string gr in grList)
                    {
                        comboGradeYear.Items.Add(gr);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            // 預設選第一項
            if (comboGradeYear.Items.Count > 0)
                comboGradeYear.SelectedIndex = 0;

            //comboGradeYear.Items.Add("3");
            //comboGradeYear.Items.Add("2");
            //comboGradeYear.Items.Add("1");

            //comboGradeYear.Text = "3";

            comboGradeYear.DropDownStyle = ComboBoxStyle.DropDownList;
            comboDept.DropDownStyle = ComboBoxStyle.DropDownList;
            comboClass.DropDownStyle = ComboBoxStyle.DropDownList;

            // 載入畢業預警欄位
            LoadDgGWColumns();

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

            ChkNotUptoGStandard.Enabled = value;
            btnExport.Enabled = btnClassReport.Enabled = value;
            tabControl1.Enabled = value;

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

                if (GradeYearClassNameDict.ContainsKey(SelectedGradeYearYear))
                {
                    foreach (string name in GradeYearClassNameDict[SelectedGradeYearYear])
                        comboClass.Items.Add(name);
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
                            ""HeaderText"": ""畢業審查"",
                            ""Name"": ""畢業審查"",
                            ""Width"": 110,
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

                //// 加入學生功能按鈕
                //DataGridViewButtonXColumn btnCol1 = new DataGridViewButtonXColumn();
                //btnCol1.Name = "學生通知單";
                //btnCol1.HeaderText = "學生通知單";
                //btnCol1.Text = "列印";
                //btnCol1.Width = 100;
                //btnCol1.UseColumnTextForButtonValue = true;
                //btnCol1.Click += BtnCol1_Click;
                //dgDataGW.Columns.Add(btnCol1);

                //// 加入班級功能按鈕
                //DataGridViewButtonXColumn btnCol2 = new DataGridViewButtonXColumn();
                //btnCol2.Name = "班級通知單";
                //btnCol2.HeaderText = "班級通知單";
                //btnCol2.Text = "列印";
                //btnCol2.Width = 100;
                //btnCol2.UseColumnTextForButtonValue = true;
                //btnCol2.Click += BtnCol2_Click;
                //dgDataGW.Columns.Add(btnCol2);

                //// 加入學生待處理
                //DataGridViewButtonXColumn btnCol3 = new DataGridViewButtonXColumn();
                //btnCol3.Name = "學生待處理";
                //btnCol3.HeaderText = "學生待處理";
                //btnCol3.Text = "加入";
                //btnCol3.Width = 80;
                //btnCol3.UseColumnTextForButtonValue = true;
                //btnCol3.Click += BtnCol3_Click;
                //dgDataGW.Columns.Add(btnCol3);

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
                },
                {
                    ""HeaderText"": ""新報部科目名稱"",
                    ""Name"": ""新報部科目名稱"",
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

                // 檢查僅顯示未達設定
                isChkNotUptoGStandard = ChkNotUptoGStandard.Checked;

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
            btnQuery.Enabled = true;
            ChkNotUptoGStandard.Visible = true;
            LoadTabDesc();
            lblMsg.Text = "共0筆";
            lnkSetReportTemplate.Visible = true;
            btnReport.Enabled = false;
            btnExport.Visible = btnClassReport.Visible = true;
            btnExport.Enabled = btnClassReport.Enabled = false;
        }

        private void tbItemChkEdit_Click(object sender, EventArgs e)
        {
            // 資料合理檢查 -- 科目級別
            SelectedTabName = ChkEditTabName;
            btnQuery.Enabled = true;
            LoadTabDesc();
            lblMsg.Text = "共" + dgDataChkEdit.Rows.Count + "筆";
            lnkSetReportTemplate.Visible = false;
            ChkNotUptoGStandard.Visible = false;
            btnExport.Visible = btnClassReport.Visible = false;
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
                msg = @"學生成績資料科目屬性合理性檢查： 
1.檢查範圍：一般及延修狀態學生科目。
2.依年級、科別、班級條件，檢查範圍內學生學期成績進行合理性檢查。以學生學期科目成績之科目名稱+級別比對學生採用之課程規劃表中的分項類別、領域、校部訂、必選修、指定學年科目名稱、課程代碼、報部科目名稱有差異的資料。
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
            LoadTabDesc();
            //lblMsg.Text = "共" + dgData2ChkEdit.Rows.Count + "筆";
            lnkSetReportTemplate.Visible = false;
            ChkNotUptoGStandard.Visible = false;
            btnExport.Visible = btnClassReport.Visible = false;
        }

        private void btnReport_Click(object sender, EventArgs e)
        {
            // 畢業預警報表
            if (SelectedTabName == GWTabName)
            {
                btnReport.Enabled = false;

                // 取得範本
                if (this.configure == null)
                {
                    FISCA.UDT.AccessHelper accessHelper = new FISCA.UDT.AccessHelper();
                    List<Configure> Configures = new List<Configure>();
                    Configures = accessHelper.Select<Configure>();
                    foreach (Configure cf in Configures)
                    {
                        if (cf.Active)
                        {
                            this.configure = cf;
                            this.configure.Decode();
                            break;
                        }
                    }
                }

                if (this.configure == null)
                {
                    MsgBox.Show("沒有設定範本");
                    btnReport.Enabled = true;
                    return;
                }
                else
                {
                    if (this.configure.Template == null)
                    {
                        MsgBox.Show("沒有設定範本");
                        btnReport.Enabled = true;
                        return;
                    }
                }

                // 處理畫面上所選學生
                // 沒選就當作所有
                SelectedReportStudentList.Clear();
                if (dgDataGW.SelectedRows.Count == 0)
                    SelectedReportStudentList = ReportStudentList;
                else
                {
                    foreach (DataGridViewRow row in dgDataGW.SelectedRows)
                    {
                        if (row.IsNewRow)
                            continue;

                        ReportStudentInfo rs = row.Tag as ReportStudentInfo;
                        if (rs != null)
                            SelectedReportStudentList.Add(rs);
                    }
                }

                // 有資料在執行
                if (SelectedReportStudentList.Count > 0)
                {
                    // 將所選的學生依照學生班級座號排序

                    try
                    {
                        SelectedReportStudentList = SelectedReportStudentList.OrderBy(x => x.ClassName).ThenBy(x =>
                        {
                            int seatNo;
                            return int.TryParse(x.SeatNo, out seatNo) ? seatNo : 0;
                        }).ToList();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }


                    bgwGrandCheckReport.RunWorkerAsync();
                }
                else
                {
                    MsgBox.Show("沒有資料。");
                    return;
                }

            }

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

            // 處理未分年級
            if (SelectedGradeYearYear == NoGradeYearStr)
            {
                comboDept.Text = "";

                // 填入班級
                if (GradeYearClassNameDict.ContainsKey(SelectedGradeYearYear))
                {
                    foreach (string name in GradeYearClassNameDict[SelectedGradeYearYear])
                    {
                        if (!comboClass.Items.Contains(name))
                            comboClass.Items.Add(name);
                    }
                    if (comboClass.Items.Count > 0)
                        comboClass.SelectedIndex = 0;
                }
            }
            else
            {
                // 一般年級
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
        }

        private void buttonLoadDesc_Click(object sender, EventArgs e)
        {
            string desc = LoadTabDesc();
            frmDesc frm = new frmDesc();
            frm.SetDesc(desc);
            frm.ShowDialog();
        }

        private void lnkSetReportTemplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ConfigForm cf = new ConfigForm();
            if (cf.ShowDialog() == DialogResult.OK)
            {
                this.configure = cf.Configure;
            }
        }

        private void ChkNotUptoGStandard_CheckedChanged(object sender, EventArgs e)
        {
            isChkNotUptoGStandard = ChkNotUptoGStandard.Checked;
            ReloadDataGridViewGW();
        }

        // 重整畢業預警審查畫面通過判斷
        private void ReloadDataGridViewGW()
        {
            dgDataGW.Rows.Clear();
            if (ReportStudentList.Count > 0)
            {
                foreach (ReportStudentInfo rs in ReportStudentList)
                {
                    // 畢業預警審查通過不顯示
                    if (isChkNotUptoGStandard && rs.GraduationCheck == "通過")
                        continue;

                    int rowIdx = dgDataGW.Rows.Add();
                    dgDataGW.Rows[rowIdx].Tag = rs;
                    dgDataGW.Rows[rowIdx].Cells["學號"].Value = rs.StudentNumber;
                    dgDataGW.Rows[rowIdx].Cells["班級"].Value = rs.ClassName;
                    int seatNo;
                    if (int.TryParse(rs.SeatNo, out seatNo))
                    {
                        dgDataGW.Rows[rowIdx].Cells["座號"].Value = seatNo;
                    }
                    else
                    {
                        dgDataGW.Rows[rowIdx].Cells["座號"].Value = 0;
                    }

                    dgDataGW.Rows[rowIdx].Cells["姓名"].Value = rs.StudentName;

                    if (rs.GraduationCheck == "通過")
                    {
                        dgDataGW.Rows[rowIdx].Cells["畢業審查"].Value = "符合畢業標準";
                    }
                    else if (rs.GraduationCheck == "不通過")
                    {
                        dgDataGW.Rows[rowIdx].Cells["畢業審查"].Value = "未達畢業標準";
                    }
                    else
                    {
                        dgDataGW.Rows[rowIdx].Cells["畢業審查"].Value = "";
                    }


                }

            }

            lblMsg.Text = "共" + dgDataGW.Rows.Count + "筆";
        }

        private void btnClassReport_Click(object sender, EventArgs e)
        {
            ControlEnable(false);
            if (ReportClassList.Count > 0 && ReportStudentList.Count > 0)
                bgwGrandCheckClassReport.RunWorkerAsync();
            else
            {
                MsgBox.Show("沒有資料。");
                ControlEnable(true);
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            ControlEnable(false);
            if (ReportStudentList.Count > 0)
                bgwGrandCheckStudentExport.RunWorkerAsync();
            else
            {
                MsgBox.Show("沒有資料。");
                ControlEnable(true);
            }
        }
    }
}
