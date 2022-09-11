using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using System.IO;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using FISCA.Authentication;

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
        /// 匯出 JSON
        /// </summary>
        public static void ExprotJSON(string ReportName, string text)
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
        public static void ExprotText(string ReportName, string text)
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
        /// 取得API學分對照
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> GetCreditMappingTable()
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
            value.Add("1", "1");
            value.Add("2", "2");
            value.Add("3", "3");
            value.Add("4", "4");
            value.Add("5", "5");
            value.Add("6", "6");
            value.Add("7", "7");
            value.Add("8", "8");
            value.Add("9", "9");
            value.Add("10", "10");
            value.Add("11", "11");
            value.Add("12", "12");

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

        /// <summary>
        /// 取得領域對照，課程代碼第19碼取2位
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> GetDomainNameMapping()
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            value.Add("00", "不分");
            value.Add("01", "語文");
            value.Add("02", "數學");
            value.Add("03", "社會");
            value.Add("04", "自然科學");
            value.Add("05", "藝術");
            value.Add("06", "綜合活動");
            value.Add("07", "科技");
            value.Add("08", "健康與體育");
            value.Add("09", "全民國防教育");
            value.Add("0A", "跨領域科目專題");
            value.Add("0B", "實作(實驗)及探索體驗");
            value.Add("0C", "職涯試探");
            value.Add("0D", "專題探究");
            value.Add("0E", "跨領域/科目統整");
            value.Add("0F", "通識性課程");
            value.Add("0G", "大學預修課程");
            value.Add("0H", "第二外國語文");
            value.Add("0I", "本土語文");
            value.Add("0J", "綜合活動與科技");
            value.Add("0R", "競技體育技能");
            value.Add("0S", "體育專業學科");
            value.Add("0T", "體育專項術科");
            value.Add("11", "特殊需求領域(身心障礙)");
            value.Add("12", "特殊需求領域(資賦優異)");
            value.Add("13", "特殊需求領域(音樂專長)");
            value.Add("14", "特殊需求領域(美術專長)");
            value.Add("15", "特殊需求領域(舞蹈專長)");
            value.Add("16", "特殊需求領域(戲劇專長)");
            value.Add("17", "特殊需求領域(體育專長)");
            value.Add("18", "特殊需求領域(實驗課程)");
            value.Add("91", "藝術才能專長");
            value.Add("92", "生涯發展");
            value.Add("93", "職能探索");
            value.Add("94", "運動防護");
            value.Add("A1", "數值控制技能");
            value.Add("A2", "精密機械製造技能");
            value.Add("A3", "模型設計與鑄造技能");
            value.Add("A4", "電腦輔助機械設計技能");
            value.Add("A5", "自動化整合技能");
            value.Add("A6", "金屬成形與管線技能");
            value.Add("B1", "車輛技能");
            value.Add("B2", "機器腳踏車技能");
            value.Add("B3", "液氣壓技能");
            value.Add("B4", "動力機械技能");
            value.Add("C1", "晶片設計技能");
            value.Add("C2", "微電腦應用技能");
            value.Add("C3", "自動控制技能");
            value.Add("C4", "電機工程技能");
            value.Add("C5", "冷凍空調技能");
            value.Add("D1", "化工及檢驗技能");
            value.Add("D2", "紡染及檢驗技能");
            value.Add("E2", "專業製圖技能");
            value.Add("E3", "土木測量技能");
            value.Add("F1", "商業與財會技能");
            value.Add("F2", "跨境商務技能");
            value.Add("F3", "資訊應用技能");
            value.Add("H1", "平面設計技能");
            value.Add("H2", "立體造形技能");
            value.Add("H3", "數位成型技能");
            value.Add("H4", "數位影音技能");
            value.Add("H5", "互動媒體技能");
            value.Add("H6", "空間設計技能");
            value.Add("I1", "農業生產與休閒生態技能");
            value.Add("I2", "動物飼養及保健技能");
            value.Add("J1", "食品加工技能");
            value.Add("J2", "檢驗分析技能");
            value.Add("K1", "整體造型技能");
            value.Add("K2", "服裝實務技能");
            value.Add("K3", "生活應用技能");
            value.Add("L1", "廚藝技能");
            value.Add("L2", "烘焙技能");
            value.Add("L3", "旅宿技能");
            value.Add("L4", "旅遊技能");
            value.Add("M1", "漁航技能");
            value.Add("M2", "漁業技能");
            value.Add("M3", "水域活動安全技能");
            value.Add("M4", "觀賞水族技能");
            value.Add("M5", "經濟水族技能");
            value.Add("M6", "區域特色水族技能");
            value.Add("M7", "海面養殖技能");
            value.Add("N1", "船舶金工技能");
            value.Add("N2", "船舶機電控制技能");
            value.Add("N3", "船舶動力技能");
            value.Add("N4", "船舶作業技能");
            value.Add("N5", "船舶操縱技能");
            value.Add("N6", "電子導航技能");
            value.Add("N7", "船舶維護與繫固作業技能");
            value.Add("O1", "視覺表現技能");
            value.Add("O2", "展演製作技能");
            value.Add("O3", "數位影音技能");
            value.Add("O4", "音樂藝術技能");
            value.Add("O5", "舞蹈藝術技能");
            value.Add("O6", "表演藝術實務技能");
            value.Add("U1", "車輛整理技能");
            value.Add("U2", "門市技能");
            value.Add("U3", "物品整理技能");
            value.Add("U4", "農園藝技能");
            value.Add("U5", "產品加工技能");
            value.Add("U6", "裝配技能");
            value.Add("U7", "生活照護技能");
            value.Add("U8", "家務處理技能");
            value.Add("U9", "餐飲製作技能");
            value.Add("UA", "旅館房務技能");
            value.Add("UB", "按摩技能");
            value.Add("UC", "紓壓保健技能");
            value.Add("UD", "民俗技能");
            value.Add("UE", "寵物照顧技能");
            value.Add("UF", "美髮技能");
            value.Add("G3", "職場實務技能");
            value.Add("G1", "英語文技能");
            value.Add("G2", "日語文技能");

            return value;
        }

        public static Dictionary<string, string> GetSchoolNMapping()
        {
            // 進校轉日校學校代碼
            Dictionary<string, string> value = new Dictionary<string, string>();
            value.Add("n.shinmin.tc.edu.tw", "191305");
            value.Add("n.mdhs.tc.edu.tw", "191309");
            value.Add("n.hzsh.tc.edu.tw", "064308");
            value.Add("n.sivs.chc.edu.tw", "070401");
            value.Add("n.ylhcvs.chc.edu.tw", "070410");
            value.Add("n.kyvs.ks.edu.tw", "121417");
            value.Add("n.klhcvs.kl.edu.tw", "171405");
            value.Add("n.sphs.hc.edu.tw", "181307");
            value.Add("n.youth.tc.edu.tw", "061316");
            value.Add("n.csvs.chc.edu.tw", "070409");
            value.Add("n2.csvs.chc.edu.tw", "070409");
            value.Add("n.tcivs.tc.edu.tw", "193407");
            value.Add("n2.tcivs.tc.edu.tw", "193407");

            return value;
        }

        public static string SubjFullName(string SubjectName, int level)
        {
            string lev = "";
            if (level == 1)
                lev = " I";

            if (level == 2)
                lev = " II";

            if (level == 3)
                lev = " III";

            if (level == 4)
                lev = " IV";

            if (level == 5)
                lev = " V";

            if (level == 6)
                lev = " VI";

            if (level == 7)
                lev = " VII";

            if (level == 8)
                lev = " VIII";

            if (level == 9)
                lev = " IX";

            if (level == 10)
                lev = " X";

            if (level == 11)
                lev = " 11";

            if (level == 12)
                lev = " 12";

            string value = SubjectName + lev;

            return value;
        }

        public static string GetCourseCodeWhereCond()
        {
            // ('進修部','實用技能學程(夜)')
            string value = " WHERE course_type NOT IN('進修部','實用技能學程(夜)') ";
            // 進校
            if (GetSchoolNMapping().Keys.Contains(DSAServices.AccessPoint))
            {
                value = " WHERE course_type IN('進修部','實用技能學程(夜)') ";
            }
            return value;
        }
    }
}
