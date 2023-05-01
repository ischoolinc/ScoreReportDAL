using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Data;

namespace SHGraduationWarning.DAO
{
    public class DataAccess
    {
        // 取得班級科別資訊
        public static List<ClassDeptInfo> GetClassDeptList()
        {
            List<ClassDeptInfo> value = new List<ClassDeptInfo>();

            try
            {
                string strSQL = @"
            SELECT
                DISTINCT dept.id AS dept_id,
                dept.name AS dept_name,
                class_name,
                class.id AS class_id
            FROM
                class
                LEFT JOIN dept ON class.ref_dept_id = dept.id
                INNER JOIN student ON class.id = student.ref_class_id
            WHERE
                student.status IN(1, 2)
            ORDER BY
                dept.name,
                class_name
            ";

                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(strSQL);
                foreach (DataRow dr in dt.Rows)
                {
                    ClassDeptInfo cd = new ClassDeptInfo();
                    cd.DeptID = dr["dept_id"] + "";
                    cd.DeptName = dr["dept_name"] + "";
                    cd.ClassID = dr["class_id"] + "";
                    cd.ClassName = dr["class_name"] + "";
                    value.Add(cd);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }

        // 取得科別id,name
        public static Dictionary<string, string> GetDeptIDNameDict()
        {
            Dictionary<string, string> value = new Dictionary<string, string>();

            try
            {
                string strSQL = @"
            SELECT
               id,
               name
            FROM
                dept            
            ORDER BY
               name;
            ";

                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(strSQL);
                foreach (DataRow dr in dt.Rows)
                {
                    string id = dr["id"] + "";
                    string name = dr["name"] + "";
                    if (!value.ContainsKey(id))
                        value.Add(id, name);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;

        }

        // 依班級年級取得學期科目級別重複  
        public static List<StudSubjectInfo> GetSemsSubjectLevelDuplicateByGradeYear(int GradeYear)
        {
            List<StudSubjectInfo> value = new List<StudSubjectInfo>();
            try
            {
                // 取得科別對照
                Dictionary<string, string> deptDict = GetDeptIDNameDict();

                QueryHelper qh = new QueryHelper();
                string strSQL = string.Format(@"
                WITH student_base AS(
                    SELECT
                        student.id AS student_id,
                        student_number,
                        seat_no,
                        student.name AS student_name,
                        class.class_name,
                        class.grade_year AS grade_year,
                        COALESCE(
                            student.ref_graduation_plan_id,
                            class.ref_graduation_plan_id
                        ) AS g_plan_id,
                        CASE
                            student.status
                            WHEN 1 THEN '一般'
                            WHEN 2 THEN '延修'
                            WHEN 4 THEN '休學'
                            WHEN 8 THEN '輟學'
                            WHEN 16 THEN '畢業或離校'
                        END AS status,
                        COALESCE(
                            student.ref_dept_id,
                            class.ref_dept_id
                        ) AS dept_id,
                        class.id AS class_id
                    FROM
                        student
                        LEFT JOIN class ON student.ref_class_id = class.id                         
                    WHERE
                        student.status <> 256
                        AND class.grade_year = {0} 
                ),
                sems_subj_score AS (
                    SELECT
                        sems_subj_score_ext.id,
                        sems_subj_score_ext.school_year,
                        sems_subj_score_ext.semester,
                        sems_subj_score_ext.grade_year,
                        sems_subj_score_ext.ref_student_id,
                        array_to_string(xpath('//Subject/@開課分項類別', subj_score_ele), '') :: text AS 分項,
                        array_to_string(xpath('//Subject/@科目', subj_score_ele), '') :: text AS 科目名稱,
                        array_to_string(xpath('//Subject/@科目級別', subj_score_ele), '') :: text AS 科目級別,
                        array_to_string(xpath('//Subject/@修課必選修', subj_score_ele), '') :: text AS 必選修,
                        array_to_string(xpath('//Subject/@修課校部訂', subj_score_ele), '') :: text AS 校部訂,
                        array_to_string(xpath('//Subject/@開課學分數', subj_score_ele), '') :: text AS 開課學分數,
                        array_to_string(xpath('//Subject/@指定學年科目名稱', subj_score_ele), '') :: text AS 指定學年科目名稱,
                        array_to_string(xpath('//Subject/@領域', subj_score_ele), '') :: text AS 領域
                    FROM
                        (
                            SELECT
                                sems_subj_score.*,
                                unnest(
                                    xpath(
                                        '//SemesterSubjectScoreInfo/Subject',
                                        xmlparse(content score_info)
                                    )
                                ) as subj_score_ele
                            FROM
                                sems_subj_score
                                INNER JOIN student_base ON sems_subj_score.ref_student_id = student_base.student_id
                        ) as sems_subj_score_ext
                ),
                student_sems_subject_2 AS (
                    SELECT
                        sems_subj_score.ref_student_id AS student_id,
                        sems_subj_score.科目名稱 AS 科目名稱,
                        sems_subj_score.科目級別 AS 科目級別,
                        COUNT(sems_subj_score.id) AS 筆數
                    FROM
                        sems_subj_score
                    GROUP BY
                        student_id,
                        科目名稱,
                        科目級別
                    HAVING
                        COUNT(sems_subj_score.id) > 1
                )
                SELECT
                    student_base.student_id AS 學生系統編號,
                    student_base.grade_year AS 年級,
                    student_base.student_number AS 學號,
                    student_base.class_name AS 班級,
                    student_base.seat_no AS 座號,
                    student_base.student_name AS 姓名,
                    student_base.g_plan_id AS g_plan_id,
                    student_base.dept_id AS dept_id,
                    student_base.class_id AS class_id,
                    sems_subj_score.school_year AS 學年度,
                    sems_subj_score.semester AS 學期,
                    sems_subj_score.grade_year AS 成績年級,
                    sems_subj_score.科目名稱,
                    sems_subj_score.科目級別,
                    sems_subj_score.領域,
                    sems_subj_score.分項,
                    sems_subj_score.必選修,
                    CASE
                        sems_subj_score.校部訂
                        WHEN '部訂' THEN '部定'
                        ELSE sems_subj_score.校部訂
                    END AS 校部訂,
                    sems_subj_score.開課學分數,
                    student_base.status AS 學生狀態,
                    sems_subj_score.id AS 學期成績系統編號,
                    sems_subj_score.指定學年科目名稱
                FROM
                    student_base
                    INNER JOIN student_sems_subject_2 ON student_base.student_id = student_sems_subject_2.student_id
                    INNER JOIN sems_subj_score ON student_sems_subject_2.student_id = sems_subj_score.ref_student_id
                    AND student_sems_subject_2.科目名稱 = sems_subj_score.科目名稱
                    AND student_sems_subject_2.科目級別 = sems_subj_score.科目級別
                ORDER BY
                    student_base.student_number,
                    sems_subj_score.科目名稱,
                    sems_subj_score.科目級別,
                    sems_subj_score.school_year,
                    sems_subj_score.semester
", GradeYear);
                DataTable dt = qh.Select(strSQL);
                foreach (DataRow dr in dt.Rows)
                {
                    StudSubjectInfo sc = new StudSubjectInfo();
                    sc.StudentID = dr["學生系統編號"] + "";
                    sc.SchoolYear = dr["學年度"] + "";
                    sc.Semester = dr["學期"] + "";
                    sc.GradeYear = dr["成績年級"] + "";
                    sc.ClassGradeYear = dr["年級"] + "";
                    sc.StudentNumber = dr["學號"] + "";
                    sc.ClassName = dr["班級"] + "";
                    sc.SeatNo = dr["座號"] + "";
                    sc.Name = dr["姓名"] + "";
                    sc.Domain = dr["領域"] + "";
                    sc.Entry = dr["分項"] + "";
                    sc.SubjectName = dr["科目名稱"] + "";
                    sc.SubjectLevel = dr["科目級別"] + "";
                    sc.SubjectLevelNew = "";
                    sc.RequiredBy = dr["校部訂"] + "";
                    sc.Required = dr["必選修"] + "";
                    sc.Credit = dr["開課學分數"] + "";
                    sc.Status = dr["學生狀態"] + "";
                    sc.SemsSubjID = dr["學期成績系統編號"] + "";
                    sc.SchoolYearSubjectName = dr["指定學年科目名稱"] + "";
                    sc.ClassID = dr["class_id"] + "";
                    sc.DeptID = dr["dept_id"] + "";
                    sc.ErrorMsgList.Add("科目級別重複");
                    // 科別名稱
                    if (deptDict.ContainsKey(sc.DeptID))
                        sc.DeptName = deptDict[sc.DeptID];

                    sc.CoursePlanID = dr["g_plan_id"] + "";
                    value.Add(sc);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return value;
        }

        // 透過課程規畫表ID，取得課程規畫表對照
        public static Dictionary<string, GPlanInfo> GetGPlanDictByIDs(List<string> ids)
        {
            Dictionary<string, GPlanInfo> value = new Dictionary<string, GPlanInfo>();

            try
            {
                if (ids == null || ids.Count == 0)
                    return value;

                string strSQL = @"
                WITH gplan_data AS(
                    SELECT
                        id,        
                        name,
                        array_to_string(xpath('//Subject/@GradeYear', subject_ele), '') :: INT AS grade_year,
                        array_to_string(xpath('//Subject/@Semester', subject_ele), '') :: INT AS semester,
                        array_to_string(xpath('//Subject/@Entry', subject_ele), '') :: text AS 分項,
                        array_to_string(xpath('//Subject/@Domain', subject_ele), '') :: text AS 領域,
                        array_to_string(xpath('//Subject/@SubjectName', subject_ele), '') :: text AS 科目,
                        array_to_string(xpath('//Subject/@Level', subject_ele), '') :: text AS 科目級別,
                        array_to_string(xpath('//Subject/@Credit', subject_ele), '') :: text AS 學分數,
                        array_to_string(xpath('//Subject/@Required', subject_ele), '') :: text AS 必選修,
                        array_to_string(xpath('//Subject/@RequiredBy', subject_ele), '') :: text AS 校部訂,
                        array_to_string(xpath('//Subject/@課程代碼', subject_ele), '') :: text AS 課程代碼,
                        array_to_string(xpath('//Subject/@指定學年科目名稱', subject_ele), '') :: text AS 指定學年科目名稱
                    FROM
                        (
                            SELECT
                                id,
                                name,                                
                                unnest(
                                    xpath(
                                        '//GraduationPlan/Subject',
                                        xmlparse(content content)
                                    )
                                ) as subject_ele
                            FROM
                                graduation_plan
                        ) AS graduation_plan
                    WHERE
                        ID IN(" + string.Join(",", ids.ToArray()) + @")
                )
                select
                    *
                from
                    gplan_data
                ORDER BY
                    name,
                    grade_year,
                    semester,
                    科目
";
                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(strSQL);
                foreach (DataRow dr in dt.Rows)
                {
                    string gpID = dr["id"] + "";
                    string gpName = dr["name"] + "";
                    if (!value.ContainsKey(gpID))
                    {
                        GPlanInfo gpi = new GPlanInfo();
                        gpi.ID = gpID;
                        gpi.Name = gpName;
                        value.Add(gpID, gpi);
                    }

                    GPlanSubjectInfo gs = new GPlanSubjectInfo();
                    gs.GradeYear = dr["grade_year"] + "";
                    gs.Semester = dr["semester"] + "";
                    gs.Entry = dr["分項"] + "";
                    gs.Domain = dr["領域"] + "";
                    gs.SubjectName = dr["科目"] + "";
                    gs.SubjectLevel = dr["科目級別"] + "";
                    gs.Required = dr["必選修"] + "";
                    gs.RequiredBy = dr["校部訂"] + "";
                    gs.Credit = dr["學分數"] + "";
                    gs.CourseCode = dr["課程代碼"] + "";
                    gs.SubjectNameByYear = dr["指定學年科目名稱"] + "";

                    //// -- 使用科目名稱與級別當key
                    // string gsKey = gs.SubjectName + "_" + gs.SubjectLevel;
                    // 使用年級+學期+科目名稱當key
                    string gsKey = gs.GradeYear + "_" + gs.Semester + "_" + gs.SubjectName ;
                    if (!value[gpID].SubjectsDict.ContainsKey(gsKey))
                        value[gpID].SubjectsDict.Add(gsKey, gs);
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
