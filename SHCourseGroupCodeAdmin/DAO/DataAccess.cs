using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using FISCA.Data;

namespace SHCourseGroupCodeAdmin.DAO
{

    public class DataAccess
    {
        Dictionary<string, string> MOEGroupCodeDict = new Dictionary<string, string>();
        Dictionary<string, string> MOEGroupNameDict = new Dictionary<string, string>();

        /// <summary>
        /// 取得課程代碼大表資料
        /// </summary>
        /// <returns></returns>
        public List<MOECourseCodeInfo> GetCourseGroupCodeList()
        {
            List<MOECourseCodeInfo> value = new List<MOECourseCodeInfo>();
            try
            {
                string query = "" +
                   " SELECT  " +
" last_update AS 最後更新日期 " +
" , group_code AS 群組代碼 " +
" , course_code AS 課程代碼 " +
" , subject_name AS 科目名稱 " +
" , entry_year AS 入學年 " +
" , (CASE require_by WHEN '部訂' THEN '部定' ELSE require_by END) AS 部定校訂 " +
" , is_required AS 必修選修 " +
" , course_type AS 課程類型 " +
" , group_type AS 群別 " +
" , subject_type AS 科別 " +
" , class_type AS 班群 " +
" , credit_period AS 授課學期學分_節數 " +
"  FROM " +
"   $moe.subjectcode ORDER BY group_code,subject_name; ";
                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    MOECourseCodeInfo data = new MOECourseCodeInfo();
                    data.group_code = dr["群組代碼"] + "";
                    data.course_code = dr["課程代碼"] + "";
                    data.subject_name = dr["科目名稱"] + "";
                    data.entry_year = dr["入學年"] + "";
                    data.require_by = dr["部定校訂"] + "";
                    data.is_required = dr["必修選修"] + "";
                    data.course_type = dr["課程類型"] + "";
                    data.group_type = dr["群別"] + "";
                    data.subject_type = dr["科別"] + "";
                    data.class_type = dr["班群"] + "";
                    data.credit_period = dr["授課學期學分_節數"] + "";

                    value.Add(data);
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }

        /// <summary>
        /// 檢查班級群科班設定
        /// </summary>
        /// <returns></returns>
        public List<ClassInfo> GetClassCourseGroup()
        {
            List<ClassInfo> value = new List<ClassInfo>();
            try
            {
                LoadMOEGroupCodeDict();


                string query = "" +
                    "SELECT " +
                    "id" +
                    ",COALESCE(grade_year,0) AS grade_year" +
                    ",class_name" +
                    ",COALESCE(gdc_code,'') AS gdc_code " +
                    " FROM class " +
                    " ORDER BY class.grade_year,display_order,class_name";
                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    ClassInfo ci = new ClassInfo();
                    ci.ClassID = dr["id"] + "";
                    ci.ClassName = dr["class_name"] + "";
                    ci.GradeYear = dr["grade_year"] + "";
                    ci.ClassGroupCode = dr["gdc_code"] + "";
                    ci.ClassGroupName = GetGroupNameByCode(ci.ClassGroupCode);
                    value.Add(ci);
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

        public List<CourseInfoChk> GetCourseCheckInfoListByGradeYear(int GradeYear)
        {
            List<CourseInfoChk> value = new List<CourseInfoChk>();

            try
            {
                QueryHelper qh = new QueryHelper();
                string query = "" +
                   " SELECT  " +
" DISTINCT course.id AS course_id " +
" , course_name " +
" , subject " +
" , subj_level " +
" , COALESCE(course.credit, null) AS c_credit " +
" , COALESCE(course.period, null) AS c_period " +
" , class.grade_year " +
" , COALESCE(course.credit,course.period) AS credit " +
" , course.school_year " +
" , course.semester " +
" , (CASE COALESCE(sc_attend.required_by,c_required_by) WHEN '1' THEN '部定' WHEN '2' THEN '校訂' ELSE '' END) AS required_by " +
" , (CASE COALESCE(sc_attend.is_required,c_is_required) WHEN '1' THEN '必修' WHEN '0' THEN '選修' ELSE '' END) AS required " +
" , COALESCE(student.gdc_code,class.gdc_code)  AS gdc_code " +
" FROM course " +
" 	INNER JOIN sc_attend " +
"  ON course.id = sc_attend.ref_course_id  " +
" 	INNER JOIN student  " +
"  ON sc_attend.ref_student_id = student.id " +
" 	INNER JOIN class " +
" 	ON student.ref_class_id = class.id " +
" WHERE  " +
"  student.status = 1 AND class.grade_year in(" + GradeYear + ") " +
" ORDER BY school_year DESC,semester DESC ";

                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    CourseInfoChk data = new CourseInfoChk();
                    data.CourseID = dr["course_id"] + "";
                    data.CourseName = dr["course_name"] + "";
                    data.SchoolYear = dr["school_year"] + "";
                    data.Semester = dr["semester"] + "";
                    data.SubjectName = dr["subject"] + "";
                    data.IsRequired = dr["required"] + "";
                    data.RequireBy = dr["required_by"] + "";
                    data.SubjectLevel = dr["subj_level"] + "";
                    data.course_code = "";
                    data.Credit = "";
                    if (dr["c_credit"] != null)
                    {
                        data.Credit = dr["c_credit"] + "";
                    }

                    data.Period = "";
                    if (dr["c_period"] != null)
                    {
                        data.Period = dr["c_period"] + "";
                    }

                    data.credit_period = "";
                    data.GroupCode = dr["gdc_code"] + "";
                    value.Add(data);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }


        /// <summary>
        /// 取得群科班代碼科目內容
        /// </summary>
        /// <returns></returns>
        public Dictionary<string, List<MOECourseCodeInfo>> GetCourseGroupCodeDict()
        {
            Dictionary<string, List<MOECourseCodeInfo>> value = new Dictionary<string, List<MOECourseCodeInfo>>();

            List<MOECourseCodeInfo> coCodeList = GetCourseGroupCodeList();
            foreach (MOECourseCodeInfo co in coCodeList)
            {
                if (!value.ContainsKey(co.group_code))
                    value.Add(co.group_code, new List<MOECourseCodeInfo>());

                value[co.group_code].Add(co);
            }
            return value;
        }
    }
}
