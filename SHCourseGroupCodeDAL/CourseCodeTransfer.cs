using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using FISCA.Data;
using System.Xml.Linq;

namespace SHCourseGroupCodeDAL
{
    public class CourseCodeTransfer
    {
        public CourseCodeTransfer() { }

        public Dictionary<string, StudentCourseCodeInfo> GetStundetCourseCodeDict(List<string> StudentIDList)
        {
            // StudentID,
            Dictionary<string, StudentCourseCodeInfo> value = new Dictionary<string, StudentCourseCodeInfo>();

            if (StudentIDList.Count > 0)
            {

                // 收集學生有課程群組代碼，找到相關科目對照
                List<string> CourseGroupCodeList = new List<string>();

                // 讀取目前學生群組代碼設定
                try
                {
                    QueryHelper qh = new QueryHelper();

                    string query = "SELECT " +
                        "student.id AS student_id" +
                        ",COALESCE(student.gdc_code,class.gdc_code) AS gdc_code " +
                        " FROM student" +
                        " LEFT OUTER JOIN class " +
                        " ON student.ref_class_id = class.id " +
                        " WHERE student.id IN(" + string.Join(",", StudentIDList.ToArray()) + ");";

                    DataTable dt = qh.Select(query);
                    foreach(DataRow dr in dt.Rows)
                    {
                        string sid = dr["student_id"] + "";
                        string code = dr["gdc_code"] + "";                       

                        StudentCourseCodeInfo scc = new StudentCourseCodeInfo();
                        scc.StudentID = sid;
                        scc.CourseGroupCode = code;

                        if (!CourseGroupCodeList.Contains(code))
                            CourseGroupCodeList.Add(code);

                        if (!value.ContainsKey(sid))
                            value.Add(sid, scc);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }                

                // 群組所屬科目代碼
                Dictionary<string, List<SubjectInfo>> SubjectCodeDict = GetSubjectCodeDictByGroupCode(CourseGroupCodeList);

                // 解析並填入學生科目
                foreach (string sid in value.Keys)
                {
                    if (SubjectCodeDict.ContainsKey(value[sid].CourseGroupCode))
                    {
                        value[sid].AddSubjectInfoList(SubjectCodeDict[value[sid].CourseGroupCode]);
                    }                 
                }
            }

            return value;
        }

        /// <summary>
        /// 透過 Group Code get subject and course code
        /// </summary>
        /// <param name="code"></param>
        /// <returns></returns>
        public Dictionary<string, List<SubjectInfo>> GetSubjectCodeDictByGroupCode(List<string> code)
        {
            Dictionary<string, List<SubjectInfo>> value = new Dictionary<string, List<SubjectInfo>>();
            // SELECT group_code,course_code,subject_name,credit_period,entry_year,require_by,is_required FROM $moe.subjectcode WHERE group_code = '108041305H11101A'


            QueryHelper qh = new QueryHelper();
            string query = "SELECT group_code,course_code,subject_name,credit_period,entry_year,require_by,is_required FROM $moe.subjectcode WHERE group_code IN('" + string.Join("''", code.ToArray()) + "')";
            try
            {
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string groupCode = dr["group_code"] + "";
                    if (!value.ContainsKey(groupCode))
                        value.Add(groupCode, new List<SubjectInfo>());

                    // 因為 ischool 資料內校部定儲存字的問題需要轉換

                    if (dr["require_by"] != null)
                    {
                        if (dr["require_by"] + "" == "部定")
                            dr["require_by"] = "部訂";
                    }


                    SubjectInfo si = new SubjectInfo(dr["subject_name"] + "", dr["entry_year"] + "", dr["require_by"] + "", dr["is_required"] + "");
                    si.SetCode(dr["group_code"] + "", dr["course_code"] + "");

                    value[groupCode].Add(si);

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


            return value;
        }


        // 取得群組代碼與說明
        public Dictionary<string, string> GetGroupCodeAndDesc()
        {
            Dictionary<string, string> value = new Dictionary<string, string>();

            // 取得班群代碼與說明
            string strSQL = "" +
" SELECT DISTINCT group_code,(entry_year|| " +
" (CASE WHEN length(course_type)>0 THEN '_'||course_type END) || " +
" (CASE WHEN length(group_type)>0 THEN '_'||group_type END) || " +
" (CASE WHEN length(subject_type)>0 THEN '_'||subject_type END) || " +
" (CASE WHEN length(class_type)>0 THEN '_'||class_type END) " +
" ) AS group_name FROM $moe.subjectcode ORDER BY group_name; ";

            try
            {
                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(strSQL);

                foreach (DataRow dr in dt.Rows)
                {
                    string code = dr["group_code"] + "";

                    if (!value.ContainsKey(code))
                        value.Add(code, dr["group_name"] + "");
                }
            }
            catch (Exception ex)
            {

            }

            return value;
        }

    }
}