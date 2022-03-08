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

        AccessHelper _accessHelper = new AccessHelper();

        List<string> StudentIDList = new List<string>();
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

            // 讀取資料
            // 取得所選學生資料
            List<StudentInfo> StudentInfoList = DataAccess.GetStudentInfoListByIDs(StudentIDList);

            //取得大表(全校3年級、6年級、12年級 學生目前身上的doc_code)


            //取得學期科目成績 (只取原始成績) 
            // 取得學期成績
            //List<rptStudSemsScoreCodeChkInfo> StudSemsScoreCodeChkInfoList = DataAccess.GetStudentSemsScoreInfo(110, 2);




            //取得課程代碼
            //計算排名百分比 (找不到的科目 或 沒有成績的科目 顯示空白)  == 比我低分的人數/總人數 *100  無條件捨去後+1 

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


            #region 產生合併欄位
            DataTable dtTable = new DataTable();
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


        }

        private void BgWorkerReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");

            //if (errorMsgList.Count > 0)
            //{
            //    MsgBox.Show("無法產生檔案。");
            //    return;
            //}


            string path = "";
            // 完成後開啟資料夾
            string p1 = "";

            //foreach (string sid in StudentDocDict.Keys)
            //{
            //    try
            //    {
            //        Document document = StudentDocDict[sid];

            //        #region 儲存檔案
            //        string reportName = "第6學期修課紀錄" + StudentDocNameDict[sid];
            //        path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports\\第6學期修課紀錄");

            //        p1 = path;
            //        if (!Directory.Exists(path))
            //            Directory.CreateDirectory(path);

            //        path = Path.Combine(path, reportName + ".docx");

            //        if (File.Exists(path))
            //        {
            //            int i = 1;
            //            while (true)
            //            {
            //                string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
            //                if (!File.Exists(newPath))
            //                {
            //                    path = newPath;
            //                    break;
            //                }
            //            }
            //        }

            //        try
            //        {
            //            document.Save(path, SaveFormat.Docx);
            //        }
            //        catch (Exception ex)
            //        {
            //            System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
            //            sd.Title = "另存新檔";
            //            sd.FileName = reportName + ".docx";
            //            sd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
            //            if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            //            {
            //                try
            //                {
            //                    document.Save(sd.FileName, Aspose.Words.SaveFormat.Docx);

            //                }
            //                catch
            //                {
            //                    FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            //                    return;
            //                }
            //            }
            //        }
            //        #endregion

            //    }
            //    catch (Exception ex)
            //    {
            //        FISCA.Presentation.Controls.MsgBox.Show("產生過程發生錯誤," + ex.Message);
            //    }
            //}

            System.Diagnostics.Process.Start(p1);
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

        private void MainForm_Load(object sender, EventArgs e)
        {
            LoadTemplate();
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
