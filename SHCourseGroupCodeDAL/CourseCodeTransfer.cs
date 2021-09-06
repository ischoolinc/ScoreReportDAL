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

        /// <summary>
        /// 捨棄不用
        /// </summary>
        /// <param name="StudGDCCodeDict"></param>
        /// <returns></returns>
        public Dictionary<string, List<StudentCourseCodeInfo>> GetStundetCourseCodeDict(List<string> StudentIDList)
        {
            Dictionary<string, List<StudentCourseCodeInfo>> value = new Dictionary<string, List<StudentCourseCodeInfo>>();
            return value;
        }

        public Dictionary<string, StudentCourseCodeInfo> GetStundetCourseCodeDictByStudCodeDict(Dictionary<string, string> StudGDCCodeDict)
        {
            // StudentID,
            Dictionary<string, StudentCourseCodeInfo> value = new Dictionary<string, StudentCourseCodeInfo>();



            // 收集學生有課程群組代碼，找到相關科目對照
            List<string> CourseGroupCodeList = new List<string>();

            // 讀取傳入學期對照取得資料，這段先註解
            //// 讀取目前學生群組代碼設定
            //try
            //{
            //    QueryHelper qh = new QueryHelper();

            //    string query = "SELECT " +
            //        "student.id AS student_id" +
            //        ",COALESCE(student.gdc_code,class.gdc_code) AS gdc_code " +
            //        " FROM student" +
            //        " LEFT OUTER JOIN class " +
            //        " ON student.ref_class_id = class.id " +
            //        " WHERE student.id IN(" + string.Join(",", StudentIDList.ToArray()) + ");";

            //    DataTable dt = qh.Select(query);
            //    foreach (DataRow dr in dt.Rows)
            //    {
            //        string sid = dr["student_id"] + "";
            //        string code = dr["gdc_code"] + "";

            //        StudentCourseCodeInfo scc = new StudentCourseCodeInfo();
            //        scc.StudentID = sid;
            //        scc.CourseGroupCode = code;

            //        if (!CourseGroupCodeList.Contains(code))
            //            CourseGroupCodeList.Add(code);

            //        if (!value.ContainsKey(sid))
            //            value.Add(sid, scc);
            //    }
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex.Message);
            //}

            foreach (string key in StudGDCCodeDict.Keys)
            {
                StudentCourseCodeInfo scc = new StudentCourseCodeInfo();
                scc.StudentID = key;
                scc.CourseGroupCode = StudGDCCodeDict[key];

                if (!value.ContainsKey(key))
                    value.Add(key, scc);
            }

            // 取得學生共通的群科班
            List<string> GroupCodeList = new List<string>();
            foreach (string val in StudGDCCodeDict.Values)
            {
                if (!string.IsNullOrWhiteSpace(val))
                {
                    if (!GroupCodeList.Contains(val))
                        GroupCodeList.Add(val);
                }
            }

            // 群組所屬科目代碼
            Dictionary<string, List<SubjectInfo>> SubjectCodeDict = GetSubjectCodeDictByGroupCode(GroupCodeList);

            // 解析並填入學生科目
            foreach (string sid in value.Keys)
            {
                if (SubjectCodeDict.ContainsKey(value[sid].CourseGroupCode))
                {
                    value[sid].AddSubjectInfoList(SubjectCodeDict[value[sid].CourseGroupCode]);
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
            string query = "" +
                "SELECT " +
                "group_code" +
                ",course_code" +
                ",subject_name" +
                ",credit_period" +
                ",entry_year" +
                ",require_by" +
                ",is_required" +
                ",course_attr" +
                " FROM $moe.subjectcode WHERE group_code IN('" + string.Join("''", code.ToArray()) + "')";
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

                    string is_required = dr["is_required"] + "";
                    string require_by = dr["require_by"] + "";
                    string course_attr = dr["course_attr"] + "";

                    // 透過課程屬性更新校部定
                    if (course_attr.Length > 0)
                    {
                        string codeC = course_attr.Substring(0, 1);
                        if (codeC == "1")
                        {
                            is_required = "必修";
                            require_by = "部訂";
                        }
                        else if (codeC == "2")
                        {
                            is_required = "必修";
                            require_by = "校訂";
                        }
                        else
                        {
                            is_required = "選修";
                            require_by = "校訂";
                        }
                    }

                    SubjectInfo si = new SubjectInfo();
                    si.SetSubjectInfo(dr["subject_name"] + "", dr["entry_year"] + "", require_by, is_required, dr["course_attr"] + "");

                    si.SetCode(dr["group_code"] + "", dr["course_code"] + "");
                    si.credit_period = dr["credit_period"] + "";
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

        public Dictionary<string, Dictionary<string, string>> GetGPlanSubjectYearSemsByGDCCode(List<string> gdcList)
        {
            Dictionary<string, Dictionary<string, string>> value = new Dictionary<string, Dictionary<string, string>>();
            try
            {
                if (gdcList.Count > 0)
                {
                    QueryHelper qh = new QueryHelper();
                    string query = "" +
                        " SELECT DISTINCT " +
    "     array_to_string(xpath('//Subject/@GradeYear', subject_ele), '')::text AS 年級 " +
    "     , array_to_string(xpath('//Subject/@Semester', subject_ele), '')::text AS 學期     " +
    "     , array_to_string(xpath('//Subject/@SubjectName', subject_ele), '')::text AS 科目 " +
    "     , array_to_string(xpath('//Subject/@Level', subject_ele), '')::text AS 科目級別 " +
    " 	,moe_group_code " +
    " FROM " +
    "     ( " +
    "         SELECT  " +
    "           " +
    "             unnest(xpath('//GraduationPlan/Subject', xmlparse(content content))) as subject_ele " +
    " 			,moe_group_code " +
    "         FROM  " +
    "             graduation_plan WHERE moe_group_code IN('" + string.Join("','", gdcList.ToArray()) + "') " +
    "     ) AS graduation_plan ";

                    DataTable dt = qh.Select(query);

                    foreach (DataRow dr in dt.Rows)
                    {
                        string code = dr["moe_group_code"] + "";
                        if (!value.ContainsKey(code))
                            value.Add(code, new Dictionary<string, string>());

                        string vv = "-1";

                        if (dr["年級"] + "" == "1" && dr["學期"] + "" == "1")
                            vv = "1";

                        if (dr["年級"] + "" == "1" && dr["學期"] + "" == "2")
                            vv = "2";

                        if (dr["年級"] + "" == "2" && dr["學期"] + "" == "1")
                            vv = "3";

                        if (dr["年級"] + "" == "2" && dr["學期"] + "" == "2")
                            vv = "4";

                        if (dr["年級"] + "" == "3" && dr["學期"] + "" == "1")
                            vv = "5";

                        if (dr["年級"] + "" == "3" && dr["學期"] + "" == "2")
                            vv = "6";


                        string key = dr["科目"] + "_" + dr["科目級別"];
                        if (!value[code].ContainsKey(key))
                            value[code].Add(key, vv);

                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return value;
        }

    }
}