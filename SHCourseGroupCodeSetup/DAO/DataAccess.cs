using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Data;
using System.Data;

namespace SHCourseGroupCodeSetup.DAO
{
    public class DataAccess
    {

        Dictionary<string, string> MOEGroupCodeDict = new Dictionary<string, string>();
        Dictionary<string, string> MOEGroupNameDict = new Dictionary<string, string>();

        /// <summary>
        /// 透過班級系統編號取得群組代碼
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public string GetClassGroupCodeByClassID(string id)
        {
            string query = "SELECT gdc_code FROM class WHERE id = " + id;
            string code = ExecuteSQLReturnString0(query);
            return code;
        }

        /// <summary>
        /// 透過學生系統編號取得群組代碼
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public string GetStudentCodeByStudentID(string id)
        {
            string query = "SELECT COALESCE(student.gdc_code,class.gdc_code) AS gdc_code FROM student LEFT JOIN class ON student.ref_class_id = class.id  WHERE student.id = " + id;
            string code = ExecuteSQLReturnString0(query);
            return code;
        }

        /// <summary>
        /// 透過班級系統編號設定群組代碼
        /// </summary>
        /// <param name="id"></param>
        /// <param name="code"></param>
        public void SetClassGroupCodeByClassID(string id, string code)
        {
            if (!string.IsNullOrWhiteSpace(id))
            {
                string query = "";
                if (string.IsNullOrWhiteSpace(code))
                {
                    query = "UPDATE class SET gdc_code = null WHERE id = " + id + " RETURNING id;";
                }
                else
                {
                    query = "UPDATE class SET gdc_code = '" + code + "' WHERE id = " + id + " RETURNING id;";
                }

                string value = ExecuteSQLReturnString0(query);
            }
        }

        public void SetClassGroupCodeByClassIDs(List<string> ids, string code)
        {
            if (ids.Count > 0)
            {
                string query = "";
                if (string.IsNullOrWhiteSpace(code))
                {
                    query = "UPDATE class SET gdc_code = NULL WHERE id IN ( " + string.Join(",", ids.ToArray()) + ") RETURNING id;";
                }
                else
                {
                    query = "UPDATE class SET gdc_code = '" + code + "' WHERE id IN ( " + string.Join(",", ids.ToArray()) + ") RETURNING id;";
                }


                string value = ExecuteSQLReturnString0(query);
            }
        }


        /// <summary>
        /// 透過學生系統編號設定群組代碼
        /// </summary>
        /// <param name="id"></param>
        /// <param name="code"></param>
        public void SetStudentGroupCodeByStudentID(string id, string code)
        {
            if (!string.IsNullOrWhiteSpace(id))
            {
                string query = "";
                if (string.IsNullOrWhiteSpace(code))
                {
                    query = "UPDATE student SET gdc_code = NULL WHERE id = " + id + " RETURNING id;";
                }
                else
                {
                    query = "UPDATE student SET gdc_code = '" + code + "' WHERE id = " + id + " RETURNING id;";
                }

                string value = ExecuteSQLReturnString0(query);
            }
        }

        public void SetStudentGroupCodeByStudentIDs(List<string> ids, string code)
        {
            if (ids.Count > 0)
            {
                string query = "";
                if (string.IsNullOrWhiteSpace(code))
                {
                    query = "UPDATE student SET gdc_code = NULL WHERE id IN ( " + string.Join(",", ids.ToArray()) + ") RETURNING id;";
                }
                else
                {
                    query = "UPDATE student SET gdc_code = '" + code + "' WHERE id IN ( " + string.Join(",", ids.ToArray()) + ") RETURNING id;";
                }

                string value = ExecuteSQLReturnString0(query);
            }
        }

        private string ExecuteSQLReturnString0(string query)
        {
            string value = "";
            try
            {
                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(query);
                if (dt != null && dt.Rows.Count > 0)
                {
                    // 取第一筆第一個值
                    value = dt.Rows[0][0] + "";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return value;
        }

        /// <summary>
        /// 取得大表群組代碼與說明
        /// </summary>
        /// <returns></returns>
        public void LoadMOEGroupCodeDict()
        {
            MOEGroupCodeDict.Clear();
            MOEGroupNameDict.Clear();

            QueryHelper qh = new QueryHelper();
            string query = "" +
" SELECT DISTINCT group_code,(entry_year|| " +
" (CASE WHEN length(course_type)>0 THEN '_'||course_type END) || " +
" (CASE WHEN length(group_type)>0 THEN '_'||group_type END) || " +
" (CASE WHEN length(subject_type)>0 THEN '_'||subject_type END) || " +
" (CASE WHEN length(class_type)>0 THEN '_'||class_type END) " +
" ) AS group_name FROM $moe.subjectcode ORDER BY group_name; ";

            DataTable dt = qh.Select(query);

            foreach (DataRow dr in dt.Rows)
            {
                string code = dr["group_code"] + "";
                string name = dr["group_name"] + "";

                if (!MOEGroupCodeDict.ContainsKey(code))
                    MOEGroupCodeDict.Add(code, name);

                if (!MOEGroupNameDict.ContainsKey(name))
                    MOEGroupNameDict.Add(name, code);
            }
        }

        public List<string> GetGroupNameList()
        {
            return MOEGroupNameDict.Keys.ToList();
        }

        public string GetGroupNameByCode(string code)
        {
            string name = "";
            if (MOEGroupCodeDict.ContainsKey(code))
                name = MOEGroupCodeDict[code];
            return name;
        }

        public string GetGroupCodeByName(string name)
        {
            string code = "";

            if (MOEGroupNameDict.ContainsKey(name))
                code = MOEGroupNameDict[name];

            return code;
        }
    }
}
