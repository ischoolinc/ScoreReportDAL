using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using FISCA.Data;
using System.Data;

namespace SHGraduationWarning
{
    public class Utility
    {
        /// <summary>
        /// 匯出 Excel
        /// </summary>
        /// <param name="inputReportName"></param>
        /// <param name="inputXls"></param>
        public static void ExportXls(string ReportName, Workbook wbXls)
        {
            string reportName = ReportName;

            string path = Path.Combine(Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".xlsx");

            Workbook wb = wbXls;

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
                wb.Save(path, SaveFormat.Xlsx);
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                SaveFileDialog sd = new SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".xlsx";
                sd.Filter = "Excel檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        wb.Save(sd.FileName, SaveFormat.Xlsx);

                    }
                    catch
                    {
                        MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
        }

        /// <summary>
        /// 匯出 JSON
        /// </summary>
        public static void ExportJSON(string ReportName, string text)
        {
            string reportName = ReportName;

            string path = Path.Combine(Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".json");



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
                StreamWriter sw = new StreamWriter(path, false, Encoding.UTF8);
                sw.WriteLine(text);
                sw.Close();
                System.Diagnostics.Process.Start("notepad.exe", path);
            }
            catch
            {
                SaveFileDialog sd = new SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".json";
                sd.Filter = "JSON檔案 (*.json)|*.json|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        StreamWriter sw = new StreamWriter(path, false, Encoding.UTF8);
                        sw.WriteLine(text);
                        sw.Close();
                    }
                    catch
                    {
                        MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
        }

        /// <summary>
        /// 匯出 Text
        /// </summary>
        public static void ExportText(string ReportName, string text)
        {
            string reportName = ReportName;

            string path = Path.Combine(Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".txt");



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
                StreamWriter sw = new StreamWriter(path, false, Encoding.UTF8);
                sw.WriteLine(text);
                sw.Close();
                System.Diagnostics.Process.Start("notepad.exe", path);
            }
            catch
            {
                SaveFileDialog sd = new SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".txt";
                sd.Filter = "Text 檔案 (*.txt)|*.txt|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        StreamWriter sw = new StreamWriter(path, false, Encoding.UTF8);
                        sw.WriteLine(text);
                        sw.Close();
                    }
                    catch
                    {
                        MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
        }

        /// <summary>
        /// 所有學生學號與狀態對應的暫存
        /// </summary>
        public static Dictionary<string, string> _AllStudentNumberStatusIDTemp = new Dictionary<string, string>();


        /// <summary>
        /// 取得所有學生學號_狀態對應 StudentNumber_Status,id
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> GetAllStudenNumberStatusDict()
        {
            Dictionary<string, string> retVal = new Dictionary<string, string>();
            QueryHelper qh = new QueryHelper();
            string strSQL = "SELECT " +
                "student.student_number" +
                ",CASE WHEN student.status = 1 THEN '一般'" +
                " WHEN student.status = 2 THEN '延修'" +
                " WHEN student.status = 4 THEN '休學'" +
                " WHEN student.status = 8 THEN '輟學'" +
                " WHEN student.status = 16 THEN '畢業或離校'" +
                " WHEN student.status = 256 THEN '刪除' ELSE '' END AS status" +
                ",student.id FROM student;";
            DataTable dt = qh.Select(strSQL);
            foreach (DataRow dr in dt.Rows)
            {
                string key = dr["student_number"] +"" + "_" + dr["status"] +"";
                
                if (!retVal.ContainsKey(key))
                    retVal.Add(key, dr["id"] +"");
            }

            return retVal;
        }

    }
}
