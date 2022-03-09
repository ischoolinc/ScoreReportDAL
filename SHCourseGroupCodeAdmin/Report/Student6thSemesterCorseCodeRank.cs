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

        int SchoolYear = 110;
        int Semester = 2;


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

            DataTable dtTable = new DataTable();
            Document docTemplate = _Configure.Template;
            if (docTemplate == null)
                docTemplate = new Document(new MemoryStream(Properties.Resources.DefaultTemplate));

            // 取得所選學生資料
            List<StudentInfo> StudentInfoList = DataAccess.GetStudentInfoListByIDs(StudentIDList, SchoolYear, Semester);

            bgWorkerReport.ReportProgress(10);

            // 取得(3,6,12年級) 110-2 學期科目原始成績
            List<rptStudSemsScoreCodeChkInfo> StudSemsScoreCodeChkInfoList = da.GetStudentSemsScoreInfo("3,6,12", SchoolYear, Semester);

            #region 產生合併欄位
            dtTable.Columns.Add("學生系統編號");
            dtTable.Columns.Add("學年度");
            dtTable.Columns.Add("學期");
            dtTable.Columns.Add("學校名稱");
            dtTable.Columns.Add("班級");
            dtTable.Columns.Add("座號");
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

            bgWorkerReport.ReportProgress(60);

            #region 群科班對照表設定
            //群科班對照表設定
            Dictionary<string, string> MappingTag1 = new Dictionary<string, string>();
            try
            {
                //< MappingConfig >
                //    < Group Name = "科班學程對照" >
                //        < Item Name = "普通科普通班" TagName = "普通科" />
                //        < Item Name = "普通科體育班班" TagName = "普通科" />
                //        < Item Name = "商經科" TagName = "商業經營科" />
                //    </ Group >
                //</ MappingConfig >

                // 解析對照設定
                XElement elmRoot = XElement.Parse(_Configure.MappingContent);
                if (elmRoot != null)
                {
                    foreach (XElement elm in elmRoot.Elements("Group"))
                    {
                        string gpName = elm.Attribute("Name").Value;
                        if (gpName == "科班學程對照")
                        {
                            foreach (XElement elm1 in elm.Elements("Item"))
                            {
                                string name = elm1.Attribute("Name").Value;
                                string tagName = elm1.Attribute("TagName").Value;
                                if (!MappingTag1.ContainsKey(name) && name.Length > 0)
                                    MappingTag1.Add(name, tagName);
                            }
                        }


                    }
                }
            }
            catch (Exception ex)
            {
                //errorMsgList.Add(ex.Message);
            }
            #endregion

            bgWorkerReport.ReportProgress(70);
            // 整理資料，填入 DataTable
            foreach (StudentInfo si in StudentInfoList)
            {
                DataRow row = dtTable.NewRow();

                row["學生系統編號"] = si.StudentID;
                row["學年度"] = SchoolYear;
                row["學期"] = Semester;
                row["學校名稱"] = K12.Data.School.ChineseName;
                row["班級"] = si.ClassName;
                row["座號"] = si.SeatNo;
                row["姓名"] = si.Name;
                row["科別"] = si.Dept;
                row["學期學業成績總平均"] = si.EntryScore;
                //學期學業成績總平均

                int index = 1;
                foreach (rptStudSemsScoreCodeChkInfo data in StudSemsScoreCodeChkInfoList)
                {
                    if (data.StudentID == si.StudentID)
                    {
                        if (index <= 60)
                        {
                            row["科目名稱" + index] = data.SubjectName;
                            row["單科學分數" + index] = data.Credit;
                            row["單科成績" + index] = data.Score;
                            row["單科成績排名百分比" + index] = data.Rank;
                            index++;
                        }
                    }
                }
                dtTable.Rows.Add(row);
            }

            bgWorkerReport.ReportProgress(80);
            docTemplate.MailMerge.Execute(dtTable);
            docTemplate.MailMerge.RemoveEmptyParagraphs = true;
            docTemplate.MailMerge.DeleteFields();

            #region Word 合併列印

            try
            {
                e.Result = new object[] { docTemplate };
            }
            catch (Exception exow)
            {
                MsgBox.Show(exow.Message);
            }

            #endregion
            bgWorkerReport.ReportProgress(100);
        }

        private void BgWorkerReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            string reportName = SchoolYear + "學年度第" + Semester + "學期學生第6學期修課紀錄";
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");

            //if (e.Error != null)
            //{
            //    MessageBox.Show(e.Error.Message);
            //    return;
            //}
            //Document doc = e.Result as Document;

            ////string reportName = _SelSchoolYear + "學年度第" + _SelSemester + "學期學期成績通知單";
            //MemoryStream memoryStream = new MemoryStream();
            //doc.Save(memoryStream, SaveFormat.Doc);
            //ePaperCloud ePaperCloud = new ePaperCloud();
            //ePaperCloud.upload_ePaper(SchoolYear, Semester, reportName, "", memoryStream, ePaperCloud.ViewerType.Student, ePaperCloud.FormatType.Docx);

            // 產生
            try
            {
                object[] objArray = (object[])e.Result;

                btnPrint.Enabled = true;

                #region 儲存檔案
                string pathQ = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
                if (!Directory.Exists(pathQ))
                    Directory.CreateDirectory(pathQ);
                string path = Path.Combine(pathQ, reportName + ".docx");
                string fPath = Path.Combine(pathQ, reportName + ".pdf");

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

                Document document = new Document();
                try
                {
                    document = (Document)objArray[0];
                    document.Save(path, SaveFormat.Docx);
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
                            document.Save(sd.FileName, Aspose.Words.SaveFormat.Docx);

                        }
                        catch (Exception ex)
                        {
                            FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗" + ex.Message, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                            return;
                        }
                    }
                }

                #region pdf
                if (File.Exists(fPath))
                {
                    int i = 1;
                    while (true)
                    {
                        string newPath = Path.GetDirectoryName(fPath) + "\\" + Path.GetFileNameWithoutExtension(fPath) + (i++) + Path.GetExtension(fPath);
                        if (!File.Exists(newPath))
                        {
                            fPath = newPath;
                            break;
                        }
                    }
                }

                Document documentPDF = new Document();
                try
                {
                    documentPDF = (Document)objArray[0];
                    documentPDF.Save(fPath, SaveFormat.Pdf);
                    System.Diagnostics.Process.Start(fPath);
                }
                catch
                {
                    System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                    sd.Title = "另存新檔";
                    sd.FileName = reportName + ".pdf";
                    sd.Filter = "PDF檔案 (*.pdf)|*.pdf|所有檔案 (*.*)|*.*";
                    if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        try
                        {
                            documentPDF.Save(sd.FileName, Aspose.Words.SaveFormat.Pdf);

                        }
                        catch (Exception ex)
                        {
                            FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗" + ex.Message, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
                #endregion


                #endregion



                MotherForm.SetStatusBarMessage(reportName + "產生完成");
            }
            catch (Exception ex)
            {
                MsgBox.Show(reportName + "產生過程發生錯誤," + ex.Message);
            }

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

            string reportName = "學生第6學期修課紀錄";

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
                catch
                {
                    MessageBox.Show("樣板開啟失敗。");
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

            bgWorkerReport.RunWorkerAsync();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
