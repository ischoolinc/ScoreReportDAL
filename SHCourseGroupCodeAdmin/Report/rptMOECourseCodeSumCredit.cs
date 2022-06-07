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
    public class rptMOECourseCodeSumCredit
    {
        BackgroundWorker _bgWorker;
        Workbook _wb;
        DataAccess da = new DataAccess();
        Dictionary<string, int> _ColIdxDict;

        public rptMOECourseCodeSumCredit()
        {
            _bgWorker = new BackgroundWorker();
            _ColIdxDict = new Dictionary<string, int>();

            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _bgWorker.WorkerReportsProgress = true;
            _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
        }

        private void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("群科班學分數總計 產生中...", e.ProgressPercentage);
        }

        private void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("群科班學分數總計 產生完成");

            if (_wb != null)
            {
                Utility.ExprotXls("群科班學分數總計", _wb);
            }
        }

        private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker.ReportProgress(1);
            // 取得資料
            List<MOECourseCodeInfo> CourseData = da.GetCourseGroupCodeList();
            Dictionary<string, MOECourseCodeSumCreditInfo> MOECourseCodeSumCreditDict = new Dictionary<string, MOECourseCodeSumCreditInfo>();
            _bgWorker.ReportProgress(50);
            // 填值到 Excel
            _wb = new Workbook(new MemoryStream(Properties.Resources.群科班學分數總計樣版));
            Worksheet wst = _wb.Worksheets[0];


            // 計算
            foreach (MOECourseCodeInfo data in CourseData)
            {
                if (!MOECourseCodeSumCreditDict.ContainsKey(data.group_code))
                {
                    MOECourseCodeSumCreditInfo da = new MOECourseCodeSumCreditInfo();
                    da.EntryYear = data.entry_year;
                    da.GroupCode = data.group_code;
                    da.GroupName = data.group_type + "_" + data.subject_type + "_" + data.class_type;
                    da.SumReqFCredit = da.SumReqTCredit = da.SumReqTotalCredit = 0;

                    MOECourseCodeSumCreditDict.Add(data.group_code, da);
                }

                char[] cp = data.credit_period.ToArray();

                foreach (char c in cp)
                {
                    decimal dc;

                    if (decimal.TryParse(c + "", out dc))
                    {
                        if (data.is_required == "必修")
                        {
                            MOECourseCodeSumCreditDict[data.group_code].SumReqTCredit += dc;
                        }
                        else
                        {
                            MOECourseCodeSumCreditDict[data.group_code].SumReqFCredit += dc;
                        }
                        MOECourseCodeSumCreditDict[data.group_code].SumReqTotalCredit += dc;
                    }else
                    {
                        Console.WriteLine(c);
                    }
                }

            }

            _ColIdxDict.Clear();

            // 讀取欄位與索引            
            for (int co = 0; co <= wst.Cells.MaxDataColumn; co++)
            {
                _ColIdxDict.Add(wst.Cells[0, co].StringValue, co);
            }

            int rowIdx = 1;

            foreach (MOECourseCodeSumCreditInfo data in MOECourseCodeSumCreditDict.Values)
            {
                wst.Cells[rowIdx, GetColIndex("群組代碼")].PutValue(data.GroupCode);
                wst.Cells[rowIdx, GetColIndex("入學年")].PutValue(data.EntryYear);
                wst.Cells[rowIdx, GetColIndex("群科班")].PutValue(data.GroupName);
                wst.Cells[rowIdx, GetColIndex("必修學分總計")].PutValue(data.SumReqTCredit);
                wst.Cells[rowIdx, GetColIndex("選修學分總計")].PutValue(data.SumReqFCredit);
                wst.Cells[rowIdx, GetColIndex("必選修學分總計")].PutValue(data.SumReqTotalCredit);

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
