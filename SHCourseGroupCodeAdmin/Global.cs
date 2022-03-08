using Aspose.Words;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseGroupCodeAdmin
{
    public class Global
    {
        public const string _UDTTableName = "ischool.SH_6thSemesterCorseCodeRank.configure";

        public static string _ProjectName = "學生第6學期修課紀錄";

        public static string _DefaultConfTypeName = "預設設定檔";

        public static void ExportMappingFieldWord()
        {

            string inputReportName = "學生第6學期修課紀錄合併欄位總表";
            string reportName = inputReportName;

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

            Document tempDoc = new Document(new MemoryStream(Properties.Resources.Template));
            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(tempDoc);
            builder.MoveToDocumentEnd();
            //builder.Writeln();
            builder.Writeln("=======合併欄位總表=======");

            builder.StartTable();
            //builder.InsertCell(); builder.Write("學年度");
            //builder.InsertCell();
            //builder.InsertField("MERGEFIELD " + "學年度" + " \\* MERGEFORMAT ", "«" + "學年度" + "»");
            //builder.EndRow();

            builder.InsertCell(); builder.Write("學校名稱");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "學校名稱" + " \\* MERGEFORMAT ", "«" + "學校名稱" + "»");
            builder.EndRow();

            builder.InsertCell(); builder.Write("班級");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "班級" + " \\* MERGEFORMAT ", "«" + "班級" + "»");
            builder.EndRow();

            builder.InsertCell(); builder.Write("座號");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "座號" + " \\* MERGEFORMAT ", "«" + "座號" + "»");
            builder.EndRow();

            builder.InsertCell(); builder.Write("姓名");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "姓名" + " \\* MERGEFORMAT ", "«" + "姓名" + "»");
            builder.EndRow();

            builder.InsertCell(); builder.Write("科別");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "科別" + " \\* MERGEFORMAT ", "«" + "科別" + "»");
            builder.EndRow();

            builder.InsertCell(); builder.Write("學期學業成績總平均");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "學期學業成績總平均" + " \\* MERGEFORMAT ", "«" + "學期學業成績總平均" + "»");
            builder.EndRow();

            builder.EndTable();

            builder.Writeln("");
            builder.Writeln("學期科目成績");
            builder.StartTable();
            builder.InsertCell(); builder.Write("科目名稱");
            builder.InsertCell(); builder.Write("單科學分數");
            builder.InsertCell(); builder.Write("單科成績");
            builder.InsertCell(); builder.Write("單科成績排名百分比");

            builder.EndRow();

            for (int i = 1; i <= 60; i++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "科目名稱" + i + " \\* MERGEFORMAT ", "«" + "N" + i + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "單科學分數" + i + " \\* MERGEFORMAT ", "«" + "C" + i + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "單科成績" + i + " \\* MERGEFORMAT ", "«" + "S" + i + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "單科成績排名百分比" + i + " \\* MERGEFORMAT ", "«" + "R" + i + "»");

                builder.EndRow();
            }


            builder.EndTable();
            builder.Writeln();


            try
            {

                tempDoc.Save(path, SaveFormat.Docx);
                System.Diagnostics.Process.Start(path);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".docx";
                sd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        tempDoc.Save(sd.FileName, SaveFormat.Docx);
                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗：" + ex.Message, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }
        }
    }
}
