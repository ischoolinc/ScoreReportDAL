using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using SHCourseGroupCodeAdmin.DAO;
using Aspose.Cells;
using System.IO;

namespace SHCourseGroupCodeAdmin.Report
{
    public class rptMOECourseCode
    {
        BackgroundWorker _bgWorker;
        Workbook _wb;
        DataAccess da = new DataAccess();
        Dictionary<string, int> _ColIdxDict;

        public rptMOECourseCode()
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
            FISCA.Presentation.MotherForm.SetStatusBarMessage("課程代碼表 產生完成");

            if (_wb != null)
            {
                Utility.ExprotXls("課程代碼表", _wb);
            }
        }

        private void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("課程代碼表 產生中...", e.ProgressPercentage);
        }

        private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker.ReportProgress(1);
            // 取得資料
            List<MOECourseCodeInfo> CourseData = da.GetCourseGroupCodeList();
            _bgWorker.ReportProgress(70);
            // 填值到 Excel
            _wb = new Workbook(new MemoryStream(Properties.Resources.課程代碼表樣版));
            Worksheet wst = _wb.Worksheets[0];

            _ColIdxDict.Clear();

            // 讀取欄位與索引            
            for (int co =0;co <=wst.Cells.MaxDataColumn; co++)
            {
                _ColIdxDict.Add(wst.Cells[0, co].StringValue, co);               
            }

            int rowIdx = 1;
            foreach(MOECourseCodeInfo data in CourseData)
            {
                wst.Cells[rowIdx, GetColIndex("群組代碼")].PutValue(data.group_code);
                wst.Cells[rowIdx, GetColIndex("課程代碼")].PutValue(data.course_code);
                wst.Cells[rowIdx, GetColIndex("科目名稱")].PutValue(data.subject_name);
                wst.Cells[rowIdx, GetColIndex("入學年")].PutValue(data.entry_year);
                wst.Cells[rowIdx, GetColIndex("部定校訂")].PutValue(data.require_by);
                wst.Cells[rowIdx, GetColIndex("必修選修")].PutValue(data.is_required);
                wst.Cells[rowIdx, GetColIndex("課程類型")].PutValue(data.course_type);
                wst.Cells[rowIdx, GetColIndex("群別")].PutValue(data.group_type);
                wst.Cells[rowIdx, GetColIndex("科別")].PutValue(data.subject_type);
                wst.Cells[rowIdx, GetColIndex("班群")].PutValue(data.class_type);
                wst.Cells[rowIdx, GetColIndex("授課學期學分/節數")].PutValue(data.credit_period);
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
