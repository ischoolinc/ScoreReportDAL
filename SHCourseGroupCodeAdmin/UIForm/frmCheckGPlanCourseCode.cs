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
                    RptGPlanCourseChkInfoPass.Add(data);
                }
            }



            _bwWorker.ReportProgress(70);
            // 輸出報表 
            _wb = new Workbook(new MemoryStream(Properties.Resources.課程規劃表開課檢核樣版));

            Worksheet wstGPChkPass = _wb.Worksheets["檢查課程規劃與採用班級群科班"];

            Worksheet wstGPChkErr = _wb.Worksheets["檢查課程規劃與採用班級群科班(有差異)"];

            Worksheet wstGPChkCoursePass = _wb.Worksheets["檢查課程課程代碼_依所屬班級"];

            Worksheet wstGPChkCourseErr = _wb.Worksheets["檢查課程課程代碼_依所屬班級(有差異)"];
            _ColIdxDict.Clear();

            // 讀取欄位與索引            
            for (int co = 0; co <= wstGPChkPass.Cells.MaxDataColumn; co++)
            {
                _ColIdxDict.Add(wstGPChkPass.Cells[0, co].StringValue, co);
            }

            int rowIdx = 1;
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
                wstGPChkErr.Cells[rowIdx, GetColIndex("說明")].PutValue(string.Join(",", data.ErrorMsgList.ToArray()));
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
