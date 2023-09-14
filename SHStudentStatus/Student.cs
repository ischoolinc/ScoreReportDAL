using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using FISCA.Data;
using System.Data;
using System.Security.AccessControl;

namespace SHStudentStatus
{
    public class Student
    {
        public static Dictionary<string, string> GetStudentStatusByStudentIDs(List<string> StudentIDs)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            // 異動代碼與身分對照
            Dictionary<string, string> UpdateCodeMapDict = new Dictionary<string, string>();
            try
            {
                if (StudentIDs == null || StudentIDs.Count == 0)
                    return dic;

                // 學生狀態，預設都一般
                List<string> StatusList = new List<string>();
                StatusList.Add("延修");
                StatusList.Add("休學");
                StatusList.Add("重讀");
                StatusList.Add("復學");
                StatusList.Add("轉科");
                StatusList.Add("畢業");

                // 取得異動代碼表
                XElement elmUpdateCodeRoot = null;
                try
                {
                    elmUpdateCodeRoot = XElement.Parse(Properties.Resources.UpdateCode_SH);
                    if (elmUpdateCodeRoot != null)
                    {
                        foreach (XElement elm in elmUpdateCodeRoot.Elements("異動"))
                        {
                            foreach (string name in StatusList)
                            {
                                if (elm.Element("原因及事項").Value.Contains(name))
                                {
                                    if (!UpdateCodeMapDict.ContainsKey(elm.Element("代號").Value))
                                        UpdateCodeMapDict.Add(elm.Element("代號").Value, name);
                                }

                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                // 取得學生最後異動
                QueryHelper qh = new QueryHelper();
                string strSQL = @"
                SELECT 
                    * 
                FROM 
                    ( 
                        SELECT 
                            ref_student_id, 
                            update_date, 
                            id, 
                            update_code, 
                            ROW_NUMBER() OVER ( 
                                PARTITION BY ref_student_id 
                                ORDER BY 
                                    ref_student_id, 
                                    update_date DESC, 
                                    id DESC 
                            ) AS row_num 
                        FROM 
                            update_record 
                        WHERE 
                            ref_student_id IN(" + string.Join(",", StudentIDs.ToArray()) + @") 
                    ) subquery 
                WHERE 
                    row_num = 1;
";

                DataTable dt = qh.Select(strSQL);

                // 整理回傳資料
                foreach (string id in StudentIDs)
                {
                    if (!dic.ContainsKey(id))
                        dic.Add(id, "一般");
                }

                foreach (DataRow dr in dt.Rows)
                {
                    string id = dr["ref_student_id"] + "";
                    string code = dr["update_code"] + "";

                    if (dic.ContainsKey(id))
                    {
                        if (UpdateCodeMapDict.ContainsKey(code))
                            dic[id] = UpdateCodeMapDict[code];
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return dic;
        }

    }
}
