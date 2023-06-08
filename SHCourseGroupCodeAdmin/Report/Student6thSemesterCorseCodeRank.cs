using Aspose.Words;
using FISCA.Presentation;
using FISCA.Presentation.Controls;
using FISCA.UDT;
using SHCourseGroupCodeAdmin.DAO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace SHCourseGroupCodeAdmin.Report
{
    public partial class Student6thSemesterCorseCodeRank : BaseForm
    {
        public Configure _Configure { get; private set; }

        BackgroundWorker bgWorkerReport = new BackgroundWorker();

        DataAccess da = new DataAccess();

        AccessHelper _accessHelper = new AccessHelper();

        List<string> StudentIDList = new List<string>();

        int SchoolYear;
        int Semester;

        Dictionary<string, Document> StudentDocDict = new Dictionary<string, Document>();
        Dictionary<string, string> StudentDocNameDict = new Dictionary<string, string>();

        /// <summary>
        /// 依照班級名稱建立資料夾再產出檔案
        /// </summary>
        bool IsAccordingToClass = false;
        Dictionary<string, string> StudentClassFolderDict = new Dictionary<string, string>();

        public Student6thSemesterCorseCodeRank()
        {
            InitializeComponent();
            bgWorkerReport.DoWork += BgWorkerReport_DoWork;
            bgWorkerReport.RunWorkerCompleted += BgWorkerReport_RunWorkerCompleted;
            bgWorkerReport.ProgressChanged += BgWorkerReport_ProgressChanged;
            bgWorkerReport.WorkerReportsProgress = true;
        }

        private void Student6thSemesterCorseCodeRank_Load(object sender, EventArgs e)
        {
            string defaultSchoolYear = K12.Data.School.DefaultSchoolYear;
            int schoolYear = 0;
            if (int.TryParse(defaultSchoolYear, out schoolYear))
                iptSchoolYear.Value = schoolYear;
            LoadTemplate();
        }
        public void SetStudentIDs(List<string> studIDs)
        {
            StudentIDList = studIDs;
        }

        private void BgWorkerReport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("第6學期修課紀錄產生中...", e.ProgressPercentage);
        }

        private void BgWorkerReport_DoWork(object sender, DoWorkEventArgs e)
        {
            bgWorkerReport.ReportProgress(1);

            StudentDocDict.Clear();
            StudentDocNameDict.Clear();
            StudentClassFolderDict.Clear();

            // 取得所選學生資料
            List<StudentInfo> StudentInfoList = DataAccess.GetStudentInfoListByIDs(StudentIDList, SchoolYear, Semester);

            bgWorkerReport.ReportProgress(10);

            // 取得(3,6,12年級) 第6學期科目原始成績
            List<rptStudSemsScoreCodeChkInfo> StudSemsScoreCodeChkInfoList = da.GetStudentSemsScoreInfo("3,6,12", SchoolYear, Semester);

            // 取得修課紀錄
            List<rptStudSemsScoreCodeChkInfo> StudSCAttendCodeInfoList = da.GetStudentCourseInfoBySchoolYearSemesterFor6thRank(SchoolYear, Semester, "3,6,12");
            //List<rptStudSemsScoreCodeChkInfo> StudSemsScoreCodeChkInfoList = da.GetStudentSemsScoreInfo("3", 109, 1);
            //List<rptStudSemsScoreCodeChkInfo> StudSCAttendCodeInfoList = da.GetStudentCourseInfoBySchoolYearSemesterFor6thRank(109, 1, "3");

            //最後要產出的學期成績&修課紀錄合併
            List<rptStudSemsScoreCodeChkInfo> ResultList = new List<rptStudSemsScoreCodeChkInfo>();

            bgWorkerReport.ReportProgress(30);

            #region 計算排名百分比
            //key=CourseCode; Value=Score
            Dictionary<string, List<decimal>> rankDic = new Dictionary<string, List<decimal>>();
            foreach (rptStudSemsScoreCodeChkInfo data in StudSemsScoreCodeChkInfoList)
            {
                if (data.CourseCode != null && data.Score.HasValue)
                    if (!rankDic.ContainsKey(data.CourseCode))
                    {
                        rankDic.Add(data.CourseCode, new List<decimal>());
                        rankDic[data.CourseCode].Add(data.Score.Value);
                    }
                    else
                    {
                        rankDic[data.CourseCode].Add(data.Score.Value);
                    }
            }

            bgWorkerReport.ReportProgress(40);

            //排序
            foreach (var k in rankDic.Keys)
            {
                var rankscores = rankDic[k];
                rankscores.Sort();
                rankscores.Reverse();
            }

            bgWorkerReport.ReportProgress(50);

            // 計算： 排名-1 / 總人數 *100  無條件捨去到整位數+1 
            foreach (var k in rankDic.Keys)
            {
                foreach (rptStudSemsScoreCodeChkInfo data in StudSemsScoreCodeChkInfoList)
                {
                    if (k == data.CourseCode && data.Score.HasValue)
                    {
                        var rankScores = rankDic[k];
                        int totalPeople = rankDic[k].Count;
                        int rankMinus1 = rankScores.IndexOf(data.Score.Value);

                        data.Rank = (rankMinus1 * 100 / totalPeople) + 1;
                    }
                }
            }

            #endregion

            bgWorkerReport.ReportProgress(40);

            #region 修課中標記處理
            //移除相同校部定、必選修、科目名稱者
            for (int i = StudSCAttendCodeInfoList.Count - 1; i >= 0; i--)
            {
                foreach (rptStudSemsScoreCodeChkInfo sem in StudSemsScoreCodeChkInfoList)
                {
                    if (StudSCAttendCodeInfoList[i].StudentID == sem.StudentID
                        && StudSCAttendCodeInfoList[i].SubjectName == sem.SubjectName
                        && StudSCAttendCodeInfoList[i].IsRequired == sem.IsRequired
                        && StudSCAttendCodeInfoList[i].RequiredBy == sem.RequiredBy)
                    {
                        StudSCAttendCodeInfoList.Remove(StudSCAttendCodeInfoList[i]);
                        //Console.WriteLine("StuID:" + StudSCAttendCodeInfoList[i].StudentID);
                        //Console.WriteLine("SubjectName:" + StudSCAttendCodeInfoList[i].SubjectName);
                        break;
                    }
                }
            }
            //合併剩餘的修課紀錄和學期成績
            ResultList.AddRange(StudSemsScoreCodeChkInfoList);
            ResultList.AddRange(StudSCAttendCodeInfoList);
            #endregion

            bgWorkerReport.ReportProgress(60);

            #region 群科班對照表設定(暫時不用)
            //群科班對照表設定
            //Dictionary<string, string> MappingTag1 = new Dictionary<string, string>();
            //try
            //{
            //    //< MappingConfig >
            //    //    < Group Name = "科班學程對照" >
            //    //        < Item Name = "普通科普通班" TagName = "普通科" />
            //    //        < Item Name = "普通科體育班班" TagName = "普通科" />
            //    //        < Item Name = "商經科" TagName = "商業經營科" />
            //    //    </ Group >
            //    //</ MappingConfig >

            //    // 解析對照設定
            //    XElement elmRoot = XElement.Parse(_Configure.MappingContent);
            //    if (elmRoot != null)
            //    {
            //        foreach (XElement elm in elmRoot.Elements("Group"))
            //        {
            //            string gpName = elm.Attribute("Name").Value;
            //            if (gpName == "科班學程對照")
            //            {
            //                foreach (XElement elm1 in elm.Elements("Item"))
            //                {
            //                    string name = elm1.Attribute("Name").Value;
            //                    string tagName = elm1.Attribute("TagName").Value;
            //                    if (!MappingTag1.ContainsKey(name) && name.Length > 0)
            //                        MappingTag1.Add(name, tagName);
            //                }
            //            }


            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    //errorMsgList.Add(ex.Message);
            //}
            #endregion
            bool skipLoop=false;
            bgWorkerReport.ReportProgress(70);
            // 整理資料，填入 DataTable
            foreach (StudentInfo si in StudentInfoList)
            {
                // 因為每位學生單檔列印，所以複製一份
                Document docTemplate = _Configure.Template.Clone();
                if (docTemplate == null)
                    docTemplate = new Document(new MemoryStream(Properties.Resources.DefaultTemplate));

                #region 產生合併欄位
                DataTable dtTable = new DataTable();
                dtTable.Columns.Add("學生系統編號");
                dtTable.Columns.Add("學年度");
                dtTable.Columns.Add("學期");
                dtTable.Columns.Add("學校名稱");
                dtTable.Columns.Add("班級");
                dtTable.Columns.Add("座號");
                dtTable.Columns.Add("學號");
                dtTable.Columns.Add("姓名");
                dtTable.Columns.Add("科別");
                dtTable.Columns.Add("學期學業成績總平均");
                for (int i = 1; i <= 60; i++)
                {
                    dtTable.Columns.Add("科目名稱" + i);
                    dtTable.Columns.Add("單科學分數" + i);
                    dtTable.Columns.Add("單科成績" + i);
                    dtTable.Columns.Add("單科成績排名百分比" + i);
                }
                #endregion

                DataRow row = dtTable.NewRow();

                row["學生系統編號"] = si.StudentID;
                row["學年度"] = SchoolYear;
                row["學期"] = Semester;
                row["學校名稱"] = K12.Data.School.ChineseName;
                row["班級"] = si.ClassName;
                row["座號"] = si.SeatNo;
                row["學號"] = si.StudentNumber;
                row["姓名"] = si.Name;
                row["科別"] = si.Dept;
                row["學期學業成績總平均"] = si.EntryScore;
                //學期學業成績總平均

                int index = 1;
                foreach (rptStudSemsScoreCodeChkInfo data in ResultList)
                {
                    if (data.StudentID == si.StudentID)
                    {
                        // Modify by Jackie Wang 20230510 增加是否列出不計學分及不須評分的判斷

                        /*if (data.NCredit == "是" && data.NScore == "是")
                            continue; //不計學分也不須評分 跳過 */
                        
                        if(data.NCredit == "是" && data.NScore == "是")
                        {
                            if (chkNCredit.Checked && chkNScore.Checked) skipLoop = false;
                            if (chkNCredit.Checked && !chkNScore.Checked) skipLoop = true;
                            if (!chkNCredit.Checked && chkNScore.Checked) skipLoop = true;
                            if (!chkNCredit.Checked && !chkNScore.Checked) skipLoop = true;

                        }
                        if (data.NCredit == "是" && data.NScore != "是")
                        {
                            if (chkNCredit.Checked && chkNScore.Checked) skipLoop = true; 
                            if (chkNCredit.Checked && !chkNScore.Checked) skipLoop = false;
                            if (!chkNCredit.Checked && chkNScore.Checked) skipLoop = true;
                            if (!chkNCredit.Checked && !chkNScore.Checked) skipLoop = true;
                        }
                        if (data.NCredit != "是" && data.NScore == "是")
                        {
                            if (chkNCredit.Checked && chkNScore.Checked) skipLoop = true;
                            if (chkNCredit.Checked && !chkNScore.Checked) skipLoop = true;
                            if (!chkNCredit.Checked && chkNScore.Checked) skipLoop = false;
                            if (!chkNCredit.Checked && !chkNScore.Checked) skipLoop = true;
                        }
                        if (data.NCredit != "是" && data.NScore != "是")
                        {
                            skipLoop = false;
                        }

                            if (skipLoop)
                        {
                           continue; //跳過
                        }                            

                        if (index <= 60)
                        {
                            row["科目名稱" + index] = data.SubjectName;
                            row["單科學分數" + index] = data.Credit;

                            if (data.IsStudying)
                            {
                                row["單科成績" + index] = "修課中";
                                row["學期學業成績總平均"] = "-";
                            }
                            else
                                row["單科成績" + index] = data.Score;

                            if (data.Rank.HasValue)
                                row["單科成績排名百分比" + index] = data.Rank + "%";

                            index++;
                        }
                    }
                }

                dtTable.Rows.Add(row);
                docTemplate.MailMerge.Execute(dtTable);
                docTemplate.MailMerge.RemoveEmptyParagraphs = true;
                docTemplate.MailMerge.DeleteFields();

                if (!StudentDocDict.ContainsKey(si.StudentID))
                {
                    StudentDocDict.Add(si.StudentID, docTemplate);
                    StudentDocNameDict.Add(si.StudentID, si.IDNumber);
                    StudentClassFolderDict.Add(si.StudentID, si.ClassName);
                }
            }

            bgWorkerReport.ReportProgress(80);

            bgWorkerReport.ReportProgress(100);
        }

        private void BgWorkerReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MotherForm.SetStatusBarMessage("完成。");
            #region 建立資料夾
            //時間戳印
            string fileDateTime = DateTime.Now.ToString("yyyyMMdd");

            //Docx
            string pathDocx = "";
            string pathDefaultFolderDocx = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports\\學生第6學期修課紀錄Docx_" + fileDateTime);
            string pathFolderDocx = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports\\學生第6學期修課紀錄Docx_" + fileDateTime);

            //PDF
            string pathPDF = "";
            string pathDefaultFolderPDF = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports\\學生第6學期修課紀錄PDF_" + fileDateTime);
            string pathFolderPDF = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports\\學生第6學期修課紀錄PDF_" + fileDateTime);
            // 完成後開啟資料夾
            string folderDocx = "";
            string folderPDF = "";

            //建立Word資料夾
            try
            {
                if (Directory.Exists(pathFolderDocx))
                {
                    int a = 1;
                    while (true)
                    {
                        string newPath = Path.GetDirectoryName(pathFolderDocx) + "\\" + Path.GetFileNameWithoutExtension(pathDefaultFolderDocx) + "_" + (a++) + Path.GetExtension(pathFolderDocx);
                        if (!Directory.Exists(newPath))
                        {
                            pathFolderDocx = newPath;
                            break;
                        }
                    }
                }
                else
                {
                    Directory.CreateDirectory(pathFolderDocx);
                }

                folderDocx = pathFolderDocx;
            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show("建立資料夾失敗：" + ex.Message);
            }

            //建立PDF資料夾
            try
            {
                if (Directory.Exists(pathFolderPDF))
                {
                    int a = 1;
                    while (true)
                    {
                        string newPath = Path.GetDirectoryName(pathFolderPDF) + "\\" + Path.GetFileNameWithoutExtension(pathDefaultFolderPDF) + "_" + (a++) + Path.GetExtension(pathFolderPDF);
                        if (!Directory.Exists(newPath))
                        {
                            pathFolderPDF = newPath;
                            break;
                        }
                    }
                }
                else
                {
                    Directory.CreateDirectory(pathFolderPDF);
                }

                folderPDF = pathFolderPDF;
            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show("建立資料夾失敗：" + ex.Message);
            }
            #endregion

            #region 單檔列印
            foreach (string sid in StudentDocDict.Keys)
            {

                #region Word
                try
                {
                    Document document = StudentDocDict[sid];

                    #region 儲存檔案
                    string reportNameSingle = StudentDocNameDict[sid];
                    string className = StudentClassFolderDict[sid];



                    pathDocx = Path.Combine(pathFolderDocx, reportNameSingle + ".docx");
                    if (IsAccordingToClass)
                        pathDocx = Path.Combine(pathFolderDocx, className, reportNameSingle + ".docx");

                    if (File.Exists(pathDocx))
                    {
                        int i = 1;
                        while (true)
                        {
                            string newPath = Path.GetDirectoryName(pathDocx) + "\\" + Path.GetFileNameWithoutExtension(pathDocx) + "_" + (i++) + Path.GetExtension(pathDocx);
                            if (!File.Exists(newPath))
                            {
                                pathDocx = newPath;
                                break;
                            }
                        }
                    }
                    try
                    {
                        document.Save(pathDocx, SaveFormat.Docx);
                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                        sd.Title = "另存新檔";
                        sd.FileName = reportNameSingle + ".docx";
                        sd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
                        if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            try
                            {
                                document.Save(sd.FileName, Aspose.Words.SaveFormat.Docx);

                            }
                            catch
                            {
                                FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                                return;
                            }
                        }
                    }
                    #endregion

                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("產生過程發生錯誤：" + ex.Message);
                }
                #endregion

                #region  PDF
                try
                {
                    Document document = StudentDocDict[sid];

                    #region 儲存檔案
                    string reportNameSingle = StudentDocNameDict[sid];
                    string className = StudentClassFolderDict[sid];

                    pathPDF = Path.Combine(pathFolderPDF, reportNameSingle + ".PDF");
                    if (IsAccordingToClass)
                        pathPDF = Path.Combine(pathFolderPDF, className, reportNameSingle + ".PDF");

                    if (File.Exists(pathPDF))
                    {
                        int i = 1;
                        while (true)
                        {
                            string newPath = Path.GetDirectoryName(pathPDF) + "\\" + Path.GetFileNameWithoutExtension(pathPDF) + "_" + (i++) + Path.GetExtension(pathPDF);
                            if (!File.Exists(newPath))
                            {
                                pathPDF = newPath;
                                break;
                            }
                        }
                    }

                    try
                    {
                        document.Save(pathPDF, SaveFormat.Pdf);
                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                        sd.Title = "另存新檔";
                        sd.FileName = reportNameSingle + ".PDF";
                        sd.Filter = "PDF檔案 (*.PDF)|*.PDF|所有檔案 (*.*)|*.*";
                        if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            try
                            {
                                document.Save(sd.FileName, Aspose.Words.SaveFormat.Pdf);

                            }
                            catch
                            {
                                FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                                return;
                            }
                        }
                    }
                    #endregion

                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("產生過程發生錯誤," + ex.Message);
                }
                #endregion
            }

            try
            {
                System.Diagnostics.Process.Start(folderDocx);
                System.Diagnostics.Process.Start(folderPDF);
            }
            catch (Exception ex)
            {
                MsgBox.Show(ex.Message);
            }
            #endregion

        }

        /// <summary>
        /// 檢視套印樣板
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lnkViewTemplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 當沒有設定檔
            if (_Configure == null) return;
            lnkViewTemplate.Enabled = false;
            #region 儲存檔案

            string reportName = "學生第6學期修課紀錄樣板";

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
                if (_Configure.Template == null)
                    _Configure.Template = new Document(new MemoryStream(Properties.Resources.DefaultTemplate));

                _Configure.Template.Save(path, Aspose.Words.SaveFormat.Docx);
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
                        _Configure.Template.Save(path, Aspose.Words.SaveFormat.Docx);
                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            lnkViewTemplate.Enabled = true;
            #endregion
        }

        /// <summary>
        /// 下載合併欄位總表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lnkViewMapColumns_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MotherForm.SetStatusBarMessage("學生第6學期修課紀錄合併欄位總表產生中...");

            // 產生合併欄位總表
            lnkViewMapColumns.Enabled = false;
            Global.ExportMappingFieldWord();
            lnkViewMapColumns.Enabled = true;
            MotherForm.SetStatusBarMessage("");
        }

        private void lnkChangeTemplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (_Configure == null) return;
            lnkChangeTemplate.Enabled = false;
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "上傳樣板";
            dialog.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    _Configure.Template = new Aspose.Words.Document(dialog.FileName);
                    List<string> fields = new List<string>(_Configure.Template.MailMerge.GetFieldNames());
                    _Configure.Encode();
                    _Configure.Save();

                }
                catch (Exception ex)
                {
                    MessageBox.Show("樣板開啟失敗：" + ex.Message);
                }
            }
            lnkChangeTemplate.Enabled = true;
        }

        /// <summary>
        /// 檢視預設樣板
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lnkDefault_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lnkDefault.Enabled = false;
            string reportName = "學生第6學期修課紀錄預設樣板";

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

            Document DefDoc = null;
            try
            {
                DefDoc = new Document(new MemoryStream(Properties.Resources.DefaultTemplate));

                DefDoc.Save(path, SaveFormat.Docx);
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
                        DefDoc.Save(path, SaveFormat.Docx);
                    }
                    catch
                    {
                        MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            lnkDefault.Enabled = true;
        }

        private void LoadTemplate()
        {
            try
            {
                UserControlEnable(false);
                List<Configure> confList = _accessHelper.Select<Configure>();
                if (confList != null && confList.Count > 0)
                {
                    _Configure = confList[0];
                    _Configure.Decode();
                }
                else
                {
                    _Configure = new Configure();
                    _Configure.Name = "學生第6學期修課";
                    _Configure.Template = new Document(new MemoryStream(Properties.Resources.DefaultTemplate));

                    _Configure.Encode();

                }
                _Configure.Save();
                UserControlEnable(true);
            }
            catch (Exception ex)
            {
                MsgBox.Show("讀取樣板發生錯誤，" + ex.Message);
            }

        }

        private void UserControlEnable(bool value)
        {
            lnkChangeTemplate.Enabled = value;
            lnkViewMapColumns.Enabled = value;
            lnkViewTemplate.Enabled = value;
            lnkDefault.Enabled = value;
            btnPrint.Enabled = value;
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            UserControlEnable(false);
            SchoolYear = iptSchoolYear.Value;
            Semester = iptSemester.Value;
            IsAccordingToClass = chkAccordingToClass.Checked;
            bgWorkerReport.RunWorkerAsync();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }


    }
}
