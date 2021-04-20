using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using SHCourseGroupCodeAdmin.DAO;
using Aspose.Cells;
using System.IO;

namespace SHCourseGroupCodeAdmin.DataCheck
{
    public class rptCheckClassGroupCode
    {
        BackgroundWorker _bgWorker;
        Workbook _wb;
        DataAccess da = new DataAccess();
        Dictionary<string, int> _ColIdxDict;

        public rptCheckClassGroupCode()
        {
            _bgWorker = new BackgroundWorker();
            _ColIdxDict = new Dictionary<string, int>();
            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _bgWorker.WorkerReportsProgress = true;
        }

        private void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("班級群科班檢查表 產生完成");

            if (_wb != null)
            {
                Utility.ExprotXls("班級群科班檢查表", _wb);
            }
        }

        private void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("班級群科班檢查表 產生中...", e.ProgressPercentage);
        }

        private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker.ReportProgress(1);
            // 取得資料
            List<ClassInfo> ClassData = da.GetClassCourseGroup();
            _bgWorker.ReportProgress(70);
            // 填值到 Excel
            _wb = new Workbook(new MemoryStream(Properties.Resources.檢查班級群科班設定樣版));
            Worksheet wst = _wb.Worksheets[0];

            _ColIdxDict.Clear();

            // 讀取欄位與索引            
            for (int co = 0; co <= wst.Cells.MaxDataColumn; co++)
            {
                _ColIdxDict.Add(wst.Cells[0, co].StringValue, co);
            }

            int rowIdx = 1;
            foreach (ClassInfo data in ClassData)
            {
                wst.Cells[rowIdx, GetColIndex("年級")].PutValue(data.GradeYear);
                wst.Cells[rowIdx, GetColIndex("班級名稱")].PutValue(data.ClassName);
                wst.Cells[rowIdx, GetColIndex("群組代碼")].PutValue(data.ClassGroupCode);
                wst.Cells[rowIdx, GetColIndex("群科班名稱")].PutValue(data.ClassGroupName);
            
                rowIdx++;
            }

            wst.AutoFitColumns();

            _bgWorker.ReportProgress(100);
        }

        private int GetColIndex(string name)
        {
            int value = 0;
            if (_ColIdxDict.ContainsKey(name))
                value = _ColIdxDict[name];

            return value;
        }

        public void Run()
        {
            _bgWorker.RunWorkerAsync();
        }
    }
}
