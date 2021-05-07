using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using System.IO;
using System.Windows.Forms;
using FISCA.Presentation.Controls;

namespace SHCourseGroupCodeAdmin
{
    public class Utility
    {
        /// <summary>
        /// 匯出 Excel
        /// </summary>
        /// <param name="inputReportName"></param>
        /// <param name="inputXls"></param>
        public static void ExprotXls(string ReportName, Workbook wbXls)
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
        /// 取得API學分對照
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string,string> GetCreditMappingTable()
        {
            /*
             學分數檢查規則：
代碼	說明
0	該學期無授課
Z	該學期有授課，但為 0 學分
n	該學期有授課之學分數(節數)
A	1學分
B	2學分
C	3學分
D	4學分
E	5學分
F	6學分
G	7學分
H	8學分
I	9學分
R	10學分
S	11學分
T	12學分
             */

            Dictionary<string, string> value = new Dictionary<string, string>();
            value.Add("0", "");
            value.Add("Z", "0");
            value.Add("n", "0");
            value.Add("A", "1");
            value.Add("B", "2");
            value.Add("C", "3");
            value.Add("D", "4");
            value.Add("E", "5");
            value.Add("F", "6");
            value.Add("G", "7");
            value.Add("H", "8");
            value.Add("I", "9");
            value.Add("R", "10");
            value.Add("S", "11");
            value.Add("T", "12");

            return value;
        }

    }
}
