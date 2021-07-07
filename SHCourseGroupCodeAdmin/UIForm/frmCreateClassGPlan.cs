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
using System.IO;
using Aspose.Cells;
using System.Xml.Linq;


namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmCreateClassGPlan : BaseForm
    {
        DataAccess _da = new DataAccess();
        string _SelectGroupName = "";
        // 已有課程規劃表
        bool hasGPlan = false;
        Workbook _wb;
        List<string> _SelGroupCodeList;
        List<string> _ErrorList;
        BackgroundWorker _bgWorker;
        BackgroundWorker _bgWorkerExport;
        Dictionary<string, int> _ColIdxDict;

        public frmCreateClassGPlan()
        {
            InitializeComponent();
            _SelGroupCodeList = new List<string>();
            _ColIdxDict = new Dictionary<string, int>();
            _ErrorList = new List<string>();
            _bgWorker = new BackgroundWorker();
            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _bgWorker.WorkerReportsProgress = true;

            _bgWorkerExport = new BackgroundWorker();
            _bgWorkerExport.DoWork += _bgWorkerExport_DoWork;
            _bgWorkerExport.ProgressChanged += _bgWorkerExport_ProgressChanged;
            _bgWorkerExport.RunWorkerCompleted += _bgWorkerExport_RunWorkerCompleted;
            _bgWorkerExport.WorkerReportsProgress = true;

        }

        private void _bgWorkerExport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
            btnExport.Enabled = true;
            if (_wb != null)
            {
                Utility.ExprotXls("班級課程規劃表資料", _wb);
            }
        }

        private void _bgWorkerExport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("班級課程規劃表 匯出中...", e.ProgressPercentage);
        }

        private void _bgWorkerExport_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorkerExport.ReportProgress(1);

            // DataTable dtAllGPlan = _da.GetAllGPlanData();
            DataTable dtAllHasMOECodeGPlan = _da.GetAllHasMOECodeGPlanData();
            _bgWorkerExport.ReportProgress(40);
            // 填值到 Excel
            _wb = new Workbook(new MemoryStream(Properties.Resources.班級課程規劃表資料樣版));

            Worksheet wst = _wb.Worksheets["課程代碼大表"];
            Worksheet wst1 = _wb.Worksheets["班級課程規劃表"];
            _ColIdxDict.Clear();

            // 讀取欄位與索引            
            for (int co = 0; co <= wst.Cells.MaxDataColumn; co++)
            {
                _ColIdxDict.Add(wst.Cells[0, co].StringValue, co);
            }

            int rowIdx = 1;

            // 整理大表資料
            Dictionary<string, rptSubjectInfo> subjectInfDict1 = new Dictionary<string, rptSubjectInfo>();
            foreach (string name in _da.GetGroupNameList())
            {
                //< Subject Category = "" Credit = "2" Domain = "不分" Entry = "學業" GradeYear = "2" Level = "4" FullName = "應用力學 IV" NotIncludedInCalc = "False" NotIncludedInCredit = "False" Required = "選修" RequiredBy = "校訂" Semester = "2" SubjectName = "應用力學" 課程代碼 = "108120401E22L0107020005" 課程類別 = "實用技能學程(日)" 開課方式 = "原班" 領域名稱 = "" 課程名稱 = "應用力學" 學分 = "2" 授課學期學分 = "002200" >
                string code = _da.GetGroupCodeByName(name);
                XElement elmRoot = _da.CourseCodeConvertToGPlanByGroupCode(code);
                foreach (XElement elm in elmRoot.Elements("Subject"))
                {
                    string key = elm.Attribute("課程代碼").Value + "_" + elm.Attribute("SubjectName").Value;

                    if (!subjectInfDict1.ContainsKey(key))
                    {
                        rptSubjectInfo subj = new rptSubjectInfo();
                        subj.MOECode = code;
                        subj.GroupName = name;
                        subj.isRequired = elm.Attribute("Required").Value;
                        subj.Domain = elm.Attribute("Domain").Value;
                        subj.CourseCode = elm.Attribute("課程代碼").Value;
                        subj.ScoreType = elm.Attribute("Entry").Value;
                        subj.SubjectName = elm.Attribute("SubjectName").Value;
                        subj.credit_period = elm.Attribute("授課學期學分").Value;
                        if (elm.Attribute("RequiredBy").Value == "部訂")
                            subj.RequiredBy = "部定";
                        else
                            subj.RequiredBy = elm.Attribute("RequiredBy").Value;

                        subjectInfDict1.Add(key, subj);
                    }
                }
            }

            // 產生大表資料
            foreach (string key in subjectInfDict1.Keys)
            {
                rptSubjectInfo subj = subjectInfDict1[key];
                wst.Cells[rowIdx, GetColIndex("群科班代碼")].PutValue(subj.MOECode);
                wst.Cells[rowIdx, GetColIndex("群科班名稱")].PutValue(subj.GroupName);
                wst.Cells[rowIdx, GetColIndex("領域")].PutValue(subj.Domain);
                wst.Cells[rowIdx, GetColIndex("分項類別")].PutValue(subj.ScoreType);
                wst.Cells[rowIdx, GetColIndex("科目名稱")].PutValue(subj.SubjectName);
                wst.Cells[rowIdx, GetColIndex("校訂部定")].PutValue(subj.RequiredBy);
                wst.Cells[rowIdx, GetColIndex("必修選修")].PutValue(subj.isRequired);
                wst.Cells[rowIdx, GetColIndex("課程代碼")].PutValue(subj.CourseCode);
                wst.Cells[rowIdx, GetColIndex("授課學期學分")].PutValue(subj.credit_period);
                rowIdx++;

            }




            //foreach (string name in _da.GetGroupNameList())
            //{
            //    //< Subject Category = "" Credit = "2" Domain = "不分" Entry = "學業" GradeYear = "2" Level = "4" FullName = "應用力學 IV" NotIncludedInCalc = "False" NotIncludedInCredit = "False" Required = "選修" RequiredBy = "校訂" Semester = "2" SubjectName = "應用力學" 課程代碼 = "108120401E22L0107020005" 課程類別 = "實用技能學程(日)" 開課方式 = "原班" 領域名稱 = "" 課程名稱 = "應用力學" 學分 = "2" 授課學期學分 = "002200" >
            //    string code = _da.GetGroupCodeByName(name);
            //    XElement elmRoot = _da.CourseCodeConvertToGPlanByGroupCode(code);
            //    foreach (XElement elm in elmRoot.Elements("Subject"))
            //    {
            //        wst.Cells[rowIdx, GetColIndex("群科班名稱")].PutValue(name);
            //        wst.Cells[rowIdx, GetColIndex("年級")].PutValue(elm.Attribute("GradeYear").Value);
            //        wst.Cells[rowIdx, GetColIndex("學期")].PutValue(elm.Attribute("Semester").Value);
            //        wst.Cells[rowIdx, GetColIndex("領域")].PutValue(elm.Attribute("Domain").Value);
            //        wst.Cells[rowIdx, GetColIndex("分項類別")].PutValue(elm.Attribute("Entry").Value);
            //        wst.Cells[rowIdx, GetColIndex("科目名稱")].PutValue(elm.Attribute("SubjectName").Value);
            //        if (elm.Attribute("RequiredBy").Value == "部訂")
            //            wst.Cells[rowIdx, GetColIndex("校訂部定")].PutValue("部定");
            //        else
            //            wst.Cells[rowIdx, GetColIndex("校訂部定")].PutValue(elm.Attribute("RequiredBy").Value);

            //        wst.Cells[rowIdx, GetColIndex("必修選修")].PutValue(elm.Attribute("Required").Value);
            //        wst.Cells[rowIdx, GetColIndex("科目級別")].PutValue(elm.Attribute("Level").Value);
            //        wst.Cells[rowIdx, GetColIndex("學分數")].PutValue(elm.Attribute("Credit").Value);
            //        wst.Cells[rowIdx, GetColIndex("課程代碼")].PutValue(elm.Attribute("課程代碼").Value);
            //        rowIdx++;
            //    }

            //}

            rowIdx = 1;

            _ColIdxDict.Clear();
            // 讀取欄位與索引            
            for (int co = 0; co <= wst1.Cells.MaxDataColumn; co++)
            {
                _ColIdxDict.Add(wst1.Cells[0, co].StringValue, co);
            }


            if (dtAllHasMOECodeGPlan != null)
            {
                // 資料整合整理
                Dictionary<string, rptSubjectInfo> subjectInfoDict2 = new Dictionary<string, rptSubjectInfo>();

                foreach (DataRow dr in dtAllHasMOECodeGPlan.Rows)
                {
                    string key = dr["課程代碼"] + "_" + dr["科目名稱"] + "";

                    if (!subjectInfoDict2.ContainsKey(key))
                    {
                        subjectInfoDict2.Add(key, new rptSubjectInfo());

                        for (int i = 0; i <= 5; i++)
                            subjectInfoDict2[key].credit[i] = "0";

                    }

                    if (dr["年級"] + "" == "1" && dr["學期"] + "" == "1")
                        subjectInfoDict2[key].credit[0] = dr["學分數"] + "";

                    if (dr["年級"] + "" == "1" && dr["學期"] + "" == "2")
                        subjectInfoDict2[key].credit[1] = dr["學分數"] + "";
                    if (dr["年級"] + "" == "2" && dr["學期"] + "" == "1")
                        subjectInfoDict2[key].credit[2] = dr["學分數"] + "";
                    if (dr["年級"] + "" == "2" && dr["學期"] + "" == "2")
                        subjectInfoDict2[key].credit[3] = dr["學分數"] + "";
                    if (dr["年級"] + "" == "3" && dr["學期"] + "" == "1")
                        subjectInfoDict2[key].credit[4] = dr["學分數"] + "";

                    if (dr["年級"] + "" == "3" && dr["學期"] + "" == "2")
                        subjectInfoDict2[key].credit[5] = dr["學分數"] + "";

                    subjectInfoDict2[key].MOECode = dr["moe_group_code"] + "";
                    subjectInfoDict2[key].GroupName = dr["name"] + "";
                    subjectInfoDict2[key].isRequired = dr["必修選修"] + "";
                    subjectInfoDict2[key].Domain = dr["領域"] + "";
                    subjectInfoDict2[key].CourseCode = dr["課程代碼"] + "";
                    subjectInfoDict2[key].ScoreType = dr["分項類別"] + "";
                    subjectInfoDict2[key].SubjectName = dr["科目名稱"] + "";
                    if (dr["校訂部定"] + "" == "部訂")
                    {
                        subjectInfoDict2[key].RequiredBy = "部定";
                    }
                    else
                    {
                        subjectInfoDict2[key].RequiredBy = dr["校訂部定"] + "";
                    }




                }

                // 產生資料至報表
                foreach (string key in subjectInfoDict2.Keys)
                {
                    rptSubjectInfo subj = subjectInfoDict2[key];
                    wst1.Cells[rowIdx, GetColIndex("群科班代碼")].PutValue(subj.MOECode);
                    wst1.Cells[rowIdx, GetColIndex("課程規劃表名稱")].PutValue(subj.GroupName);
                    wst1.Cells[rowIdx, GetColIndex("領域")].PutValue(subj.Domain);
                    wst1.Cells[rowIdx, GetColIndex("分項類別")].PutValue(subj.ScoreType);
                    wst1.Cells[rowIdx, GetColIndex("科目名稱")].PutValue(subj.SubjectName);
                    wst1.Cells[rowIdx, GetColIndex("校訂部定")].PutValue(subj.RequiredBy);
                    wst1.Cells[rowIdx, GetColIndex("必修選修")].PutValue(subj.isRequired);
                    wst1.Cells[rowIdx, GetColIndex("課程代碼")].PutValue(subj.CourseCode);
                    wst1.Cells[rowIdx, GetColIndex("授課學期學分")].PutValue(string.Join("",subj.credit.ToArray()));
                    rowIdx++;
                }

            }

            wst.AutoFitColumns();
            wst1.AutoFitColumns();
            _bgWorkerExport.ReportProgress(100);
        }

        private int GetColIndex(string name)
        {
            int value = 0;
            if (_ColIdxDict.ContainsKey(name))
                value = _ColIdxDict[name];

            return value;
        }

        private void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("建立班級課程規劃表 產生完成");

            if (_ErrorList.Count == 0)
            {
                MsgBox.Show("建立完成");
            }

            btnCreate.Enabled = true;

        }

        private void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("建立班級課程規劃表 產生中...", e.ProgressPercentage);
        }

        private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker.ReportProgress(1);
            _ErrorList.Clear();
            // 有傳入 Group Code 再執行
            if (_SelGroupCodeList.Count > 0)
            {
                foreach (string gpCode in _SelGroupCodeList)
                {

                    string errMsg = _da.WriteToGPlanByGroupCode(gpCode);
                    if (!string.IsNullOrEmpty(errMsg))
                        _ErrorList.Add(errMsg);
                    _bgWorker.ReportProgress(50);
                }
            }
            _bgWorker.ReportProgress(100);
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            // 取得畫面上選擇群科班名稱
            _SelectGroupName = cbxGroupName.Text;

            hasGPlan = false;
            string gpName = _da.GetGPlanNameByGroupName(_SelectGroupName);

            bool hasGPName = _da.CheckHasGPlan(gpName);
            if (hasGPName)
            {
                hasGPlan = true;
                MsgBox.Show(gpName + " 班級課程規劃表內已存在，無法建立。");
                return;
                //MsgBox.Show(gpName + " 班級課程規劃表內已存在，按「是」新增科目與更新學分數，請問是否執行？", "資料已存在", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            }

            _SelGroupCodeList.Clear();
            string gpCode = _da.GetGroupCodeByName(_SelectGroupName);
            _SelGroupCodeList.Add(gpCode);
            btnCreate.Enabled = false;
            _bgWorker.RunWorkerAsync();

        }

        private void frmCreateClassGPlan_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            cbxGroupName.Enabled = false;
            btnCreate.Enabled = false;

            // 取得課程代碼大表資料並填入群科班選項
            _da.LoadMOEGroupCodeDict();

            foreach (string name in _da.GetGroupNameList())
            {
                cbxGroupName.Items.Add(name);
            }

            if (cbxGroupName.Items.Count > 0)
            {
                cbxGroupName.SelectedIndex = 0;
                btnCreate.Enabled = true;
            }

            cbxGroupName.Enabled = true;

        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            btnExport.Enabled = false;
            _bgWorkerExport.RunWorkerAsync();

        }
    }
}
