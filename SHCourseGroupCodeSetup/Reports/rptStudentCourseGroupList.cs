using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using System.IO;
using FISCA.Data;
using System.ComponentModel;
using System.Data;
using SHCourseGroupCodeSetup.DAO;


namespace SHCourseGroupCodeSetup.Reports
{
    public class rptStudentCourseGroupList
    {
        List<string> studentIDList = new List<string>();
        BackgroundWorker _bgWorker;
        Workbook _wb;
        DataAccess da = new DataAccess();

        public rptStudentCourseGroupList(List<string> ids)
        {
            studentIDList = ids;
            _bgWorker = new BackgroundWorker();
            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _bgWorker.WorkerReportsProgress = true;
        }

        private void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("學生群科班清單 產生完成");

            if (_wb != null)
            {
                Utility.ExprotXls("學生群科班清單", _wb);
            }
        }

        private void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("學生群科班清單 產生中...", e.ProgressPercentage);
        }

        private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker.ReportProgress(1);
            da.LoadMOEGroupCodeDict();
            _bgWorker.ReportProgress(10);
            string query = "SELECT " +
                "student.id AS student_id" +
                ",student_number" +
                ",class.class_name" +
                ",student.name AS student_name" +
                ",student.seat_no" +
                ",COALESCE(student.gdc_code,class.gdc_code) AS gdc_code" +
                " FROM student " +
                "LEFT JOIN class " +
                "ON student.ref_class_id = class.id " +
                "WHERE student.id IN(" + string.Join(",", studentIDList.ToArray()) + ") " +
                "ORDER BY student_number,class_name,seat_no;";

            DataTable dt = null;
            try
            {
                QueryHelper qh = new QueryHelper();
                dt = qh.Select(query);
            }
            catch (Exception ex)
            {

            }

            _bgWorker.ReportProgress(70);
            // 填值到 Excel
            _wb = new Workbook(new MemoryStream(Properties.Resources.高中學生群科班清單樣板));

            int rowIdx = 1;

            if (dt != null)
            {
                foreach(DataRow dr in dt.Rows)
                {
                    _wb.Worksheets[0].Cells[rowIdx, 0].PutValue(dr["student_id"] +"");
                    _wb.Worksheets[0].Cells[rowIdx, 1].PutValue(dr["student_number"] + "");
                    _wb.Worksheets[0].Cells[rowIdx, 2].PutValue(dr["class_name"] + "");
                    _wb.Worksheets[0].Cells[rowIdx, 3].PutValue(dr["seat_no"] + "");
                    _wb.Worksheets[0].Cells[rowIdx, 4].PutValue(dr["student_name"] + "");
                    string code = dr["gdc_code"] + "";
                    _wb.Worksheets[0].Cells[rowIdx, 5].PutValue(code);
                    _wb.Worksheets[0].Cells[rowIdx, 6].PutValue(da.GetGroupNameByCode(code));
                
                    rowIdx++;
                }
            }

            _bgWorker.ReportProgress(100);
        }

        public void Run()
        {
            _bgWorker.RunWorkerAsync();
        }
    }
}
