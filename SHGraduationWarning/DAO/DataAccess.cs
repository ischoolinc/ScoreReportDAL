using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using FISCA.Data;
using FISCA.LogAgent;

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
                dept.id AS dept_id,
                dept.name AS dept_name,
                class_name,
                class.id AS class_id,
                class.grade_year
            FROM
                class
                LEFT JOIN dept ON class.ref_dept_id = dept.id
                INNER JOIN student ON class.id = student.ref_class_id
            WHERE
                student.status IN(1, 2)
            GROUP BY
                dept_id,
                dept_name,
                class_name,
                class_id
            ORDER BY
                class.grade_year DESC,
                class.display_order,
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
                    cd.GradeYear = dr["grade_year"] + "";
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

        // 依取得學期科目內容  
        public static List<StudSubjectInfo> GetSemsSubjectInfo(string DeptID, string ClassID, string textName)
        {


            List<StudSubjectInfo> value = new List<StudSubjectInfo>();
            try
            {
                // 取得科別對照
                Dictionary<string, string> deptDict = GetDeptIDNameDict();

                string condition = " ";

                if (!string.IsNullOrEmpty(DeptID))
                    condition = " dept_id = " + DeptID + "";

                if (!string.IsNullOrEmpty(ClassID))
                    condition = " class_id = " + ClassID + "";

                if (!string.IsNullOrEmpty(textName) && !string.IsNullOrEmpty(DeptID))
                    condition += " AND student_name LIKE '" + textName + "%'";
                else
                {
                    if (string.IsNullOrEmpty(DeptID))
                        condition += " student_name LIKE '" + textName + "%'";
                }



                QueryHelper qh = new QueryHelper();
                string strSQL = string.Format(@"
                WITH student_base_source AS(
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
                    student.status IN(1,2,4)
            ),
            student_base AS (
                SELECT
                    *
                FROM
                    student_base_source
                WHERE
                    {0}
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
                INNER JOIN sems_subj_score ON student_base.student_id = sems_subj_score.ref_student_id
    
            ORDER BY
                student_base.student_number,
                sems_subj_score.科目名稱,
                sems_subj_score.科目級別,
                sems_subj_score.school_year,
                sems_subj_score.semester
", condition);
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


        // 比對課程規畫表SQL(學期科目成績為主比對課規)
        public static List<StudSubjectInfo> GetSemsSubjectLevelCheckGraduationPlan1(string GradeYear, string DeptID, string ClassID)
        {
            List<StudSubjectInfo> value = new List<StudSubjectInfo>();
            try
            {
                // 取得科別對照
                Dictionary<string, string> deptDict = GetDeptIDNameDict();

                string condition = @"
                    SELECT
                        " + GradeYear + @"::INT AS grade_year, -- NULL時為全部年級 
                        NULL :: INTEGER AS class_id,
                        NULL :: INTEGER AS dept_id
                ";


                // SELECT 3::INT AS grade_year -- NULL時為全部年級
                //	           , NULL::TEXT AS dept_name--NULL時為全部科別

                if (!string.IsNullOrEmpty(DeptID))
                {
                    condition = @"
                    SELECT                        
                        NULL :: INTEGER AS class_id,
                        " + DeptID + " AS dept_id," +
                        "" + GradeYear + " AS grade_year ";
                }


                if (!string.IsNullOrEmpty(ClassID))
                {
                    condition = @"
                    SELECT                        
                        " + ClassID + " AS class_id," +
                        "" + DeptID + " AS dept_id," +
                    "" + GradeYear + " AS grade_year ";
                }

                QueryHelper qh = new QueryHelper();
                string strSQL = string.Format(@"
                 WITH row AS(
	               {0}
                ),
                target_student AS(
	                SELECT
                        student.id AS student_id,
                        graduation_plan.id AS graduation_plan_id,
                        graduation_plan.name AS graduation_plan_name,
		                class.id AS class_id,
		                class.class_name,
		                dept.id AS dept_id,
		                student.student_number,
		                student.seat_no,
		                student.name AS student_name,
                        dept.name AS dept_name
	                FROM
		                row
		                INNER JOIN class
			                       ON (
                                   class.grade_year = row.grade_year 
				                   AND ( 
                                        row.class_id is null 
                                        OR class.id = row.class_id
                                    )   
			                    )
		                INNER JOIN student
			                ON student.ref_class_id = class.id
			                AND student.status IN (1, 2)
		                INNER JOIN dept
			                ON dept.id = COALESCE(student.ref_dept_id, class.ref_dept_id)
			                AND (
				                row.dept_id IS NULL
				                 OR dept.id = row.dept_id
			                )
                        INNER JOIN graduation_plan 
                            ON graduation_plan.id = COALESCE(
                                student.ref_graduation_plan_id,
                                class.ref_graduation_plan_id
                            )
                ),
                target_student_with_sems_history AS(
                    SELECT
                        ref_student_id AS student_id,
                        graduation_plan_id,
                        grade_year :: SMALLINT,
                        semester :: SMALLINT,
                        MAX(school_year) AS school_year
                    FROM
                        sems_subj_score
                        INNER JOIN target_student
			                ON target_student.student_id = sems_subj_score.ref_student_id
                    GROUP BY
                        ref_student_id,
                        graduation_plan_id,
                        grade_year,
                        semester
                ),
                graduation_plan_expand AS(
                    SELECT
                        graduation_plan_expand.graduation_plan_id,
                        array_to_string(xpath('//Subject/@GradeYear', subject_ele), '') :: SMALLINT AS grade_year,
                        array_to_string(xpath('//Subject/@Semester', subject_ele), '') :: SMALLINT AS semester,
                        array_to_string(xpath('//Subject/@SubjectName', subject_ele), '') :: TEXT AS subject_name,
                        array_to_string(xpath('//Subject/@Level', subject_ele), '') :: TEXT AS subject_level,
                        array_to_string(xpath('//Subject/@Domain', subject_ele), '') :: TEXT AS domain,
                        array_to_string(xpath('//Subject/@分組名稱', subject_ele), '') :: TEXT AS 分組名稱,
                        (
                            '0' || array_to_string(xpath('//Subject/@分組修課學分數', subject_ele), '')
                        ) :: INTEGER AS 分組修課學分數
                    FROM
                        (
                            SELECT
                                target_graduation.graduation_plan_id,
                                unnest(
                                    xpath(
                                        '//GraduationPlan/Subject',
                                        xmlparse(content graduation_plan.content)
                                    )
                                ) AS subject_ele
                            FROM
                                (
					                SELECT
						                DISTINCT
						                graduation_plan_id
					                FROM
						                target_student
				                ) AS target_graduation
                                INNER JOIN graduation_plan 
					                ON graduation_plan.id = target_graduation.graduation_plan_id
                        ) AS graduation_plan_expand
                ),
                graduation_plan_expand_with_student AS(
                    SELECT
                        target_student_with_sems_history.student_id,
		                graduation_plan_expand.*
                    FROM
                        graduation_plan_expand
                    INNER JOIN target_student_with_sems_history 
		                ON target_student_with_sems_history.graduation_plan_id = graduation_plan_expand.graduation_plan_id
		                AND target_student_with_sems_history.grade_year = graduation_plan_expand.grade_year
		                AND target_student_with_sems_history.semester = graduation_plan_expand.semester
    
                ),
                subject_expand AS(
                    SELECT
                        sems_subj_score_ext.sems_subj_score_id,
                        sems_subj_score_ext.student_id,
                        sems_subj_score_ext.graduation_plan_id,
                        sems_subj_score_ext.grade_year,
                        sems_subj_score_ext.school_year,
                        sems_subj_score_ext.semester,
                        array_to_string(xpath('//Subject/@科目', subj_score_ele), '') :: text AS subject_name,
                        array_to_string(xpath('//Subject/@科目級別', subj_score_ele), '') :: text AS subject_level,
                        (
                            '0' || array_to_string(xpath('//Subject/@開課學分數', subj_score_ele), '')
                        ) :: INTEGER AS credit,
                        array_to_string(xpath('//Subject/@不計學分', subj_score_ele), '')::text AS 不計學分,
                        array_to_string(xpath('//Subject/@不需評分', subj_score_ele), '')::text AS 不需評分
                    FROM
                        (
                            SELECT
                                sems_subj_score.id AS sems_subj_score_id,
                                sems_subj_score.ref_student_id AS student_id,
                                target_student_with_sems_history.graduation_plan_id,
                                sems_subj_score.grade_year,
                                sems_subj_score.school_year,
                                sems_subj_score.semester,
                                unnest(
                                    xpath(
                                        '//SemesterSubjectScoreInfo/Subject',
                                        xmlparse(content score_info)
                                    )
                                ) as subj_score_ele
                            FROM
                                sems_subj_score
                                INNER JOIN target_student_with_sems_history 
					                ON target_student_with_sems_history.student_id = sems_subj_score.ref_student_id
					                AND target_student_with_sems_history.grade_year = sems_subj_score.grade_year
					                AND target_student_with_sems_history.school_year = sems_subj_score.school_year
					                AND target_student_with_sems_history.semester = sems_subj_score.semester
                        ) as sems_subj_score_ext
                ),
                target_data AS(
                    SELECT
                        subject_expand.sems_subj_score_id,
                        COALESCE(subject_expand.student_id, graduation_plan_expand_with_student.student_id) AS student_id,
                        COALESCE(subject_expand.graduation_plan_id, graduation_plan_expand_with_student.graduation_plan_id) AS graduation_plan_id,
                        COALESCE(subject_expand.grade_year, graduation_plan_expand_with_student.grade_year) AS grade_year,
                        subject_expand.school_year AS school_year,
                        COALESCE(subject_expand.semester, graduation_plan_expand_with_student.semester) AS semester,
                        COALESCE(subject_expand.subject_name, graduation_plan_expand_with_student.subject_name) AS subject_name,
                        COALESCE(subject_expand.subject_level, graduation_plan_expand_with_student.subject_level) AS subject_level,
                        subject_expand.credit,
                        subject_expand.不計學分,
                        subject_expand.不需評分,
                        graduation_plan_expand_with_student.domain,
                        graduation_plan_expand_with_student.分組名稱,
                        graduation_plan_expand_with_student.分組修課學分數,
		                suggest_graduation_plan.subject_name AS suggest_subject_name,
		                suggest_graduation_plan.subject_level AS suggest_subject_level
                    FROM
                        subject_expand
                    FULL OUTER JOIN graduation_plan_expand_with_student
                        ON subject_expand.student_id = graduation_plan_expand_with_student.student_id
                        AND subject_expand.grade_year = graduation_plan_expand_with_student.grade_year
                        AND subject_expand.semester = graduation_plan_expand_with_student.semester
                        AND subject_expand.subject_name = graduation_plan_expand_with_student.subject_name
                        AND subject_expand.subject_level = graduation_plan_expand_with_student.subject_level
                    FULL OUTER JOIN graduation_plan_expand_with_student AS suggest_graduation_plan
                        ON subject_expand.student_id = suggest_graduation_plan.student_id
                        AND subject_expand.grade_year = suggest_graduation_plan.grade_year
                        AND subject_expand.semester = suggest_graduation_plan.semester
                        AND subject_expand.subject_name = suggest_graduation_plan.subject_name
                ),
                target_match AS (
                    --成績年級、學期、科目、級別比對到的資料
                    SELECT
                        target_data.student_id,
                        target_data.grade_year,
                        target_data.semester,
                        target_data.school_year,
                        target_data.subject_name,
                        target_data.subject_level,
                        target_data.分組名稱,
                        target_data.分組修課學分數,
                        target_data.credit
                    FROM
                        target_data
                    WHERE
                        target_data.student_id IS NOT NULL
                        AND target_data.sems_subj_score_id IS NOT NULL
                        AND target_data.分組名稱 IS NOT NULL
                    ORDER BY
                        student_id,
                        grade_year,
                        semester,
                        subject_name,
                        subject_level
                ),
                target_mismatch AS (
                    --成績年級、學期、科目、級別比對到的資料
                    SELECT
                        target_data.student_id,
                        target_data.grade_year,
                        target_data.semester,
                        target_data.school_year,
                        target_data.subject_name,
                        target_data.subject_level,
                        target_data.credit,
                        target_data.不計學分,
                        target_data.不需評分
                    FROM
                        target_data
                    WHERE
		                target_data.student_id IS NOT NULL
                        AND target_data.分組名稱 IS NULL
                    ORDER BY
                        student_id,
                        grade_year,
                        semester,
                        subject_name,
                        subject_level
                ),
                graduation_plan_mismatch_all AS(
                    SELECT
                        target_data.student_id,
                        target_data.graduation_plan_id,
                        target_data.grade_year,
                        target_data.semester,
                        target_data.domain,
                        target_data.subject_name,
                        target_data.subject_level,
                        target_data.分組名稱,
                        target_data.分組修課學分數
                    FROM
                        target_data
                    WHERE
		                target_data.student_id IS NOT NULL
                        AND target_data.sems_subj_score_id IS NULL
                ),
                graduation_plan_mismatch AS (
                    --學生的課程規劃表有，卻沒有比對到的成績年級、學期、科目、級別  
                    SELECT
                        graduation_plan_mismatch_all.*
                    FROM
                        graduation_plan_mismatch_all
                    WHERE
                        分組名稱 = ''
                ),
                graduation_plan_subject_group_mismatch AS (
                    --學生課程規劃表中，規劃該年級的課程群組中，在比對到的資料中學分總數不符合的
                    /*
                     -- 條件
                     找出課程群組中的課程，用科目名稱+級別比對實際修課或成績的學分數是否符合群組設定的學分
                     */
                    SELECT
                        target_data.student_id,
                        target_data.graduation_plan_id,
                        target_data.grade_year,
                        target_data.semester,
                        target_data.分組名稱,
                        target_data.分組修課學分數,
                        SUM(target_data.credit) AS sum_credit
                    FROM
                        target_data
                    WHERE
                        target_data.student_id IS NOT NULL
                        AND target_data.分組名稱 <> ''
                    GROUP BY
                        target_data.student_id,
                        target_data.graduation_plan_id,
                        target_data.grade_year,
                        target_data.semester,
                        target_data.分組名稱,
                        target_data.分組修課學分數
                    HAVING
                        SUM(target_data.credit) <> target_data.分組修課學分數
                )
                -- sheet1
                SELECT
                    target_data.student_id AS 學生系統編號,
                    target_student.student_number AS 學號,
                    target_student.dept_name AS 科別名稱,
                    target_student.class_name AS 班級,
                    target_student.seat_no AS 座號,
                    target_student.student_name AS 姓名,
                    target_student.graduation_plan_name AS 使用課程規劃表,
                    target_data.school_year AS 學年度,
                    target_data.semester AS 學期,
                    target_data.grade_year AS 成績年級,
                    target_data.subject_name AS 科目名稱,
                    target_data.subject_level AS 科目級別,
                    target_data.credit AS 學分數,
                    target_data.不計學分,
                    target_data.不需評分,
	                target_data.suggest_subject_name AS 新科目名稱,
                    target_data.suggest_subject_level AS 新科目級別
                FROM
                    target_data
                    INNER JOIN target_student
                        ON target_data.student_id = target_student.student_id
                WHERE
	                target_data.student_id IS NOT NULL
                    AND target_data.分組名稱 IS NULL
                ORDER BY
                    班級,
                    座號,
                    學號,
                    學年度,
                    學期,
                    科目名稱

", condition);

                //    Utility.ExportText("sql1", strSQL);

                DataTable dt = qh.Select(strSQL);
                foreach (DataRow dr in dt.Rows)
                {
                    StudSubjectInfo sc = new StudSubjectInfo();
                    sc.StudentID = dr["學生系統編號"] + "";
                    sc.SchoolYear = dr["學年度"] + "";
                    sc.Semester = dr["學期"] + "";
                    sc.GradeYear = dr["成績年級"] + "";
                    sc.StudentNumber = dr["學號"] + "";
                    sc.ClassName = dr["班級"] + "";
                    sc.SeatNo = dr["座號"] + "";
                    sc.Name = dr["姓名"] + "";
                    sc.SubjectName = dr["科目名稱"] + "";
                    sc.SubjectLevel = dr["科目級別"] + "";
                    sc.GPName = dr["使用課程規劃表"] + "";
                    sc.SubjectNameNew = dr["新科目名稱"] + "";
                    sc.SubjectLevelNew = dr["新科目級別"] + "";
                    sc.DeptName = dr["科別名稱"] + "";
                    sc.Credit = dr["學分數"] + "";
                    sc.NotIncludedInCredit = dr["不計學分"] + "";
                    sc.NotIncludedInCalc = dr["不需評分"] + "";

                    value.Add(sc);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }

        // 比對課程規畫表SQL(課規為主比對學期科目成績)
        public static List<DataRow> GetSemsSubjectLevelCheckGraduationPlan2(string GradeYear, string DeptID, string ClassID)
        {
            List<DataRow> value = new List<DataRow>();
            try
            {
                // 取得科別對照
                Dictionary<string, string> deptDict = GetDeptIDNameDict();

                string condition = @"
                    SELECT
                        " + GradeYear + @"::INT AS grade_year, -- NULL時為全部年級 
                        NULL :: INTEGER AS class_id,
                        NULL :: INTEGER AS dept_id
                ";


                // SELECT 3::INT AS grade_year -- NULL時為全部年級
                //	           , NULL::TEXT AS dept_name--NULL時為全部科別

                if (!string.IsNullOrEmpty(DeptID))
                {
                    condition = @"
                    SELECT                        
                        NULL :: INTEGER AS class_id,
                        " + DeptID + " AS dept_id," +
                        "" + GradeYear + " AS grade_year ";
                }


                if (!string.IsNullOrEmpty(ClassID))
                {
                    condition = @"
                    SELECT                        
                        " + ClassID + " AS class_id," +
                        "" + DeptID + " AS dept_id," +
                    "" + GradeYear + " AS grade_year ";
                }

                QueryHelper qh = new QueryHelper();
                string strSQL = string.Format(@"
                WITH row AS(
	               {0}
                ),
                target_student AS(
	                SELECT
                        student.id AS student_id,
                        graduation_plan.id AS graduation_plan_id,
                        graduation_plan.name AS graduation_plan_name,
		                class.id AS class_id,
		                class.class_name,
		                dept.id AS dept_id,
		                student.student_number,
		                student.seat_no,
		                student.name AS student_name,
                        dept.name AS dept_name
	                FROM
		                row
		                INNER JOIN class
			                       ON (
                                   class.grade_year = row.grade_year 
				                   AND ( 
                                        row.class_id is null 
                                        OR class.id = row.class_id
                                    )   
			                    )
		                INNER JOIN student
			                ON student.ref_class_id = class.id
			                AND student.status IN (1, 2)
		                INNER JOIN dept
			                ON dept.id = COALESCE(student.ref_dept_id, class.ref_dept_id)
			                AND (
				                row.dept_id IS NULL
				                 OR dept.id = row.dept_id
			                )
                        INNER JOIN graduation_plan 
                            ON graduation_plan.id = COALESCE(
                                student.ref_graduation_plan_id,
                                class.ref_graduation_plan_id
                            )
                ),
                target_student_with_sems_history AS(
                    SELECT
                        ref_student_id AS student_id,
                        graduation_plan_id,
                        grade_year :: SMALLINT,
                        semester :: SMALLINT,
                        MAX(school_year) AS school_year
                    FROM
                        sems_subj_score
                        INNER JOIN target_student
			                ON target_student.student_id = sems_subj_score.ref_student_id
                    GROUP BY
                        ref_student_id,
                        graduation_plan_id,
                        grade_year,
                        semester
                ),
                graduation_plan_expand AS(
                    SELECT
                        graduation_plan_expand.graduation_plan_id,
                        array_to_string(xpath('//Subject/@GradeYear', subject_ele), '') :: SMALLINT AS grade_year,
                        array_to_string(xpath('//Subject/@Semester', subject_ele), '') :: SMALLINT AS semester,
                        array_to_string(xpath('//Subject/@SubjectName', subject_ele), '') :: TEXT AS subject_name,
                        array_to_string(xpath('//Subject/@Level', subject_ele), '') :: TEXT AS subject_level,
                        array_to_string(xpath('//Subject/@Domain', subject_ele), '') :: TEXT AS domain,
                        array_to_string(xpath('//Subject/@分組名稱', subject_ele), '') :: TEXT AS 分組名稱,
                        (
                            '0' || array_to_string(xpath('//Subject/@分組修課學分數', subject_ele), '')
                        ) :: INTEGER AS 分組修課學分數,
                        array_to_string(xpath('//Subject/@Entry', subject_ele), '') :: TEXT AS 分項類別,
                        array_to_string(xpath('//Subject/@RequiredBy', subject_ele), '') :: TEXT AS 校部訂,
                        array_to_string(xpath('//Subject/@Required', subject_ele), '') :: TEXT AS 必選修
                    FROM
                        (
                            SELECT
                                target_graduation.graduation_plan_id,
                                unnest(
                                    xpath(
                                        '//GraduationPlan/Subject',
                                        xmlparse(content graduation_plan.content)
                                    )
                                ) AS subject_ele
                            FROM
                                (
					                SELECT
						                DISTINCT
						                graduation_plan_id
					                FROM
						                target_student
				                ) AS target_graduation
                                INNER JOIN graduation_plan 
					                ON graduation_plan.id = target_graduation.graduation_plan_id
                        ) AS graduation_plan_expand
                ),
                graduation_plan_expand_with_student AS(
                    SELECT
                        target_student_with_sems_history.student_id,
		                graduation_plan_expand.*
                    FROM
                        graduation_plan_expand
                    INNER JOIN target_student_with_sems_history 
		                ON target_student_with_sems_history.graduation_plan_id = graduation_plan_expand.graduation_plan_id
		                AND target_student_with_sems_history.grade_year = graduation_plan_expand.grade_year
		                AND target_student_with_sems_history.semester = graduation_plan_expand.semester
    
                ),
                subject_expand AS(
                    SELECT
                        sems_subj_score_ext.sems_subj_score_id,
                        sems_subj_score_ext.student_id,
                        sems_subj_score_ext.graduation_plan_id,
                        sems_subj_score_ext.grade_year,
                        sems_subj_score_ext.school_year,
                        sems_subj_score_ext.semester,
                        array_to_string(xpath('//Subject/@科目', subj_score_ele), '') :: text AS subject_name,
                        array_to_string(xpath('//Subject/@科目級別', subj_score_ele), '') :: text AS subject_level,
                        (
                            '0' || array_to_string(xpath('//Subject/@開課學分數', subj_score_ele), '')
                        ) :: INTEGER AS credit,
                        array_to_string(xpath('//Subject/@不計學分', subj_score_ele), '')::text AS 不計學分,
                        array_to_string(xpath('//Subject/@不需評分', subj_score_ele), '')::text AS 不需評分
                    FROM
                        (
                            SELECT
                                sems_subj_score.id AS sems_subj_score_id,
                                sems_subj_score.ref_student_id AS student_id,
                                target_student_with_sems_history.graduation_plan_id,
                                sems_subj_score.grade_year,
                                sems_subj_score.school_year,
                                sems_subj_score.semester,
                                unnest(
                                    xpath(
                                        '//SemesterSubjectScoreInfo/Subject',
                                        xmlparse(content score_info)
                                    )
                                ) as subj_score_ele
                            FROM
                                sems_subj_score
                                INNER JOIN target_student_with_sems_history 
					                ON target_student_with_sems_history.student_id = sems_subj_score.ref_student_id
					                AND target_student_with_sems_history.grade_year = sems_subj_score.grade_year
					                AND target_student_with_sems_history.school_year = sems_subj_score.school_year
					                AND target_student_with_sems_history.semester = sems_subj_score.semester
                        ) as sems_subj_score_ext
                ),
                target_data AS(
                    SELECT
                        subject_expand.sems_subj_score_id,
                        COALESCE(subject_expand.student_id, graduation_plan_expand_with_student.student_id) AS student_id,
                        COALESCE(subject_expand.graduation_plan_id, graduation_plan_expand_with_student.graduation_plan_id) AS graduation_plan_id,
                        COALESCE(subject_expand.grade_year, graduation_plan_expand_with_student.grade_year) AS grade_year,
                        subject_expand.school_year AS school_year,
                        COALESCE(subject_expand.semester, graduation_plan_expand_with_student.semester) AS semester,
                        COALESCE(subject_expand.subject_name, graduation_plan_expand_with_student.subject_name) AS subject_name,
                        COALESCE(subject_expand.subject_level, graduation_plan_expand_with_student.subject_level) AS subject_level,
                        subject_expand.credit,
                        subject_expand.不計學分,
                        subject_expand.不需評分,
                        graduation_plan_expand_with_student.domain,
                        graduation_plan_expand_with_student.分組名稱,
                        graduation_plan_expand_with_student.分組修課學分數,
                        graduation_plan_expand_with_student.分項類別 AS 新分項類別,
                        graduation_plan_expand_with_student.必選修 AS 新必選修,
                        graduation_plan_expand_with_student.校部訂 AS 新校部訂,
		                suggest_graduation_plan.subject_name AS suggest_subject_name,
		                suggest_graduation_plan.subject_level AS suggest_subject_level
                    FROM
                        subject_expand
                    FULL OUTER JOIN graduation_plan_expand_with_student
                        ON subject_expand.student_id = graduation_plan_expand_with_student.student_id
                        AND subject_expand.grade_year = graduation_plan_expand_with_student.grade_year
                        AND subject_expand.semester = graduation_plan_expand_with_student.semester
                        AND subject_expand.subject_name = graduation_plan_expand_with_student.subject_name
                        AND subject_expand.subject_level = graduation_plan_expand_with_student.subject_level
                    FULL OUTER JOIN graduation_plan_expand_with_student AS suggest_graduation_plan
                        ON subject_expand.student_id = suggest_graduation_plan.student_id
                        AND subject_expand.grade_year = suggest_graduation_plan.grade_year
                        AND subject_expand.semester = suggest_graduation_plan.semester
                        AND subject_expand.subject_name = suggest_graduation_plan.subject_name
                ),
                target_match AS (
                    --成績年級、學期、科目、級別比對到的資料
                    SELECT
                        target_data.student_id,
                        target_data.grade_year,
                        target_data.semester,
                        target_data.school_year,
                        target_data.subject_name,
                        target_data.subject_level,
                        target_data.分組名稱,
                        target_data.分組修課學分數,
                        target_data.credit
                    FROM
                        target_data
                    WHERE
                        target_data.student_id IS NOT NULL
                        AND target_data.sems_subj_score_id IS NOT NULL
                        AND target_data.分組名稱 IS NOT NULL
                    ORDER BY
                        student_id,
                        grade_year,
                        semester,
                        subject_name,
                        subject_level
                ),
                target_mismatch AS (
                    --成績年級、學期、科目、級別比對到的資料
                    SELECT
                        target_data.student_id,
                        target_data.grade_year,
                        target_data.semester,
                        target_data.school_year,
                        target_data.subject_name,
                        target_data.subject_level,
                        target_data.credit,
                        target_data.不計學分,
                        target_data.不需評分
                    FROM
                        target_data
                    WHERE
		                target_data.student_id IS NOT NULL
                        AND target_data.分組名稱 IS NULL
                    ORDER BY
                        student_id,
                        grade_year,
                        semester,
                        subject_name,
                        subject_level
                ),
                graduation_plan_mismatch_all AS(
                    SELECT
                        target_data.student_id,
                        target_data.graduation_plan_id,
                        target_data.grade_year,
                        target_data.semester,
                        target_data.domain,
                        target_data.subject_name,
                        target_data.subject_level,
                        target_data.分組名稱,
                        target_data.分組修課學分數,
                        target_data.新分項類別,
                        target_data.新必選修,
                        target_data.新校部訂 
                    FROM
                        target_data
                    WHERE
		                target_data.student_id IS NOT NULL
                        AND target_data.sems_subj_score_id IS NULL
                ),
                graduation_plan_mismatch AS (
                    --學生的課程規劃表有，卻沒有比對到的成績年級、學期、科目、級別  
                    SELECT
                        graduation_plan_mismatch_all.*
                    FROM
                        graduation_plan_mismatch_all
                    WHERE
                        分組名稱 = ''
                ),
                graduation_plan_subject_group_mismatch AS (
                    --學生課程規劃表中，規劃該年級的課程群組中，在比對到的資料中學分總數不符合的
                    /*
                     -- 條件
                     找出課程群組中的課程，用科目名稱+級別比對實際修課或成績的學分數是否符合群組設定的學分
                     */
                    SELECT
                        target_data.student_id,
                        target_data.graduation_plan_id,
                        target_data.grade_year,
                        target_data.semester,
                        target_data.分組名稱,
                        target_data.分組修課學分數,
                        SUM(target_data.credit) AS sum_credit
                    FROM
                        target_data
                    WHERE
                        target_data.student_id IS NOT NULL
                        AND target_data.分組名稱 <> ''
                    GROUP BY
                        target_data.student_id,
                        target_data.graduation_plan_id,
                        target_data.grade_year,
                        target_data.semester,
                        target_data.分組名稱,
                        target_data.分組修課學分數
                    HAVING
                        SUM(target_data.credit) <> target_data.分組修課學分數
                )
               -- sheet2
                SELECT
                    target_student.student_id AS 學生系統編號,
                    target_student.student_number AS 學號,
                    target_student.dept_name AS 科別名稱,
                    target_student.class_name AS 班級,
                    target_student.seat_no AS 座號,
                    target_student.student_name AS 姓名,
                    target_student.graduation_plan_name AS 使用課程規劃表,
                    graduation_plan_mismatch.grade_year AS 成績年級,
                    graduation_plan_mismatch.semester AS 學期,
                    graduation_plan_mismatch.domain AS 領域,
                    graduation_plan_mismatch.subject_name AS 科目名稱,
                    graduation_plan_mismatch.subject_level AS 科目級別,
                    graduation_plan_mismatch.新分項類別,
                    graduation_plan_mismatch.新必選修,
                    graduation_plan_mismatch.新校部訂
                FROM
                    graduation_plan_mismatch
                    INNER JOIN target_student
		                ON graduation_plan_mismatch.student_id = target_student.student_id
                ORDER BY
                    班級,
                    座號,
                    學號,
                    CASE WHEN graduation_plan_mismatch.domain = '特殊需求領域(身心障礙)' THEN 1 ELSE 0 END,
                    成績年級,
                    學期,
                    科目名稱
", condition);

                DataTable dt = qh.Select(strSQL);
                foreach (DataRow dr in dt.Rows)
                {
                    value.Add(dr);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }

        // 比對課程規畫表SQL(課規為主檢查課程群組學分總數不符合)
        public static List<DataRow> GetSemsSubjectLevelCheckGraduationPlan3(string GradeYear, string DeptID, string ClassID)
        {
            List<DataRow> value = new List<DataRow>();
            try
            {
                // 取得科別對照
                Dictionary<string, string> deptDict = GetDeptIDNameDict();

                string condition = @"
                    SELECT
                        " + GradeYear + @"::INT AS grade_year, -- NULL時為全部年級 
                        NULL :: INTEGER AS class_id,
                        NULL :: INTEGER AS dept_id
                ";


                // SELECT 3::INT AS grade_year -- NULL時為全部年級
                //	           , NULL::TEXT AS dept_name--NULL時為全部科別

                if (!string.IsNullOrEmpty(DeptID))
                {
                    condition = @"
                    SELECT                        
                        NULL :: INTEGER AS class_id,
                        " + DeptID + " AS dept_id," +
                        "" + GradeYear + " AS grade_year ";
                }


                if (!string.IsNullOrEmpty(ClassID))
                {
                    condition = @"
                    SELECT                        
                        " + ClassID + " AS class_id," +
                        "" + DeptID + " AS dept_id," +
                    "" + GradeYear + " AS grade_year ";
                }

                QueryHelper qh = new QueryHelper();
                string strSQL = string.Format(@"
                WITH row AS(
	               {0}
                ),
                target_student AS(
	                SELECT
                        student.id AS student_id,
                        graduation_plan.id AS graduation_plan_id,
                        graduation_plan.name AS graduation_plan_name,
		                class.id AS class_id,
		                class.class_name,
		                dept.id AS dept_id,
		                student.student_number,
		                student.seat_no,
		                student.name AS student_name,
                        dept.name AS dept_name
	                FROM
		                row
		                INNER JOIN class
			                       ON (
                                   class.grade_year = row.grade_year 
				                   AND ( 
                                        row.class_id is null 
                                        OR class.id = row.class_id
                                    )   
			                    )
		                INNER JOIN student
			                ON student.ref_class_id = class.id
			                AND student.status IN (1, 2)
		                INNER JOIN dept
			                ON dept.id = COALESCE(student.ref_dept_id, class.ref_dept_id)
			                AND (
				                row.dept_id IS NULL
				                 OR dept.id = row.dept_id
			                )
                        INNER JOIN graduation_plan 
                            ON graduation_plan.id = COALESCE(
                                student.ref_graduation_plan_id,
                                class.ref_graduation_plan_id
                            )
                ),
                target_student_with_sems_history AS(
                    SELECT
                        ref_student_id AS student_id,
                        graduation_plan_id,
                        grade_year :: SMALLINT,
                        semester :: SMALLINT,
                        MAX(school_year) AS school_year
                    FROM
                        sems_subj_score
                        INNER JOIN target_student
			                ON target_student.student_id = sems_subj_score.ref_student_id
                    GROUP BY
                        ref_student_id,
                        graduation_plan_id,
                        grade_year,
                        semester
                ),
                graduation_plan_expand AS(
                    SELECT
                        graduation_plan_expand.graduation_plan_id,
                        array_to_string(xpath('//Subject/@GradeYear', subject_ele), '') :: SMALLINT AS grade_year,
                        array_to_string(xpath('//Subject/@Semester', subject_ele), '') :: SMALLINT AS semester,
                        array_to_string(xpath('//Subject/@SubjectName', subject_ele), '') :: TEXT AS subject_name,
                        array_to_string(xpath('//Subject/@Level', subject_ele), '') :: TEXT AS subject_level,
                        array_to_string(xpath('//Subject/@Domain', subject_ele), '') :: TEXT AS domain,
                        array_to_string(xpath('//Subject/@分組名稱', subject_ele), '') :: TEXT AS 分組名稱,
                        (
                            '0' || array_to_string(xpath('//Subject/@分組修課學分數', subject_ele), '')
                        ) :: INTEGER AS 分組修課學分數
                    FROM
                        (
                            SELECT
                                target_graduation.graduation_plan_id,
                                unnest(
                                    xpath(
                                        '//GraduationPlan/Subject',
                                        xmlparse(content graduation_plan.content)
                                    )
                                ) AS subject_ele
                            FROM
                                (
					                SELECT
						                DISTINCT
						                graduation_plan_id
					                FROM
						                target_student
				                ) AS target_graduation
                                INNER JOIN graduation_plan 
					                ON graduation_plan.id = target_graduation.graduation_plan_id
                        ) AS graduation_plan_expand
                ),
                graduation_plan_expand_with_student AS(
                    SELECT
                        target_student_with_sems_history.student_id,
		                graduation_plan_expand.*
                    FROM
                        graduation_plan_expand
                    INNER JOIN target_student_with_sems_history 
		                ON target_student_with_sems_history.graduation_plan_id = graduation_plan_expand.graduation_plan_id
		                AND target_student_with_sems_history.grade_year = graduation_plan_expand.grade_year
		                AND target_student_with_sems_history.semester = graduation_plan_expand.semester
    
                ),
                subject_expand AS(
                    SELECT
                        sems_subj_score_ext.sems_subj_score_id,
                        sems_subj_score_ext.student_id,
                        sems_subj_score_ext.graduation_plan_id,
                        sems_subj_score_ext.grade_year,
                        sems_subj_score_ext.school_year,
                        sems_subj_score_ext.semester,
                        array_to_string(xpath('//Subject/@科目', subj_score_ele), '') :: text AS subject_name,
                        array_to_string(xpath('//Subject/@科目級別', subj_score_ele), '') :: text AS subject_level,
                        (
                            '0' || array_to_string(xpath('//Subject/@開課學分數', subj_score_ele), '')
                        ) :: INTEGER AS credit,
                        array_to_string(xpath('//Subject/@不計學分', subj_score_ele), '')::text AS 不計學分,
                        array_to_string(xpath('//Subject/@不需評分', subj_score_ele), '')::text AS 不需評分
                    FROM
                        (
                            SELECT
                                sems_subj_score.id AS sems_subj_score_id,
                                sems_subj_score.ref_student_id AS student_id,
                                target_student_with_sems_history.graduation_plan_id,
                                sems_subj_score.grade_year,
                                sems_subj_score.school_year,
                                sems_subj_score.semester,
                                unnest(
                                    xpath(
                                        '//SemesterSubjectScoreInfo/Subject',
                                        xmlparse(content score_info)
                                    )
                                ) as subj_score_ele
                            FROM
                                sems_subj_score
                                INNER JOIN target_student_with_sems_history 
					                ON target_student_with_sems_history.student_id = sems_subj_score.ref_student_id
					                AND target_student_with_sems_history.grade_year = sems_subj_score.grade_year
					                AND target_student_with_sems_history.school_year = sems_subj_score.school_year
					                AND target_student_with_sems_history.semester = sems_subj_score.semester
                        ) as sems_subj_score_ext
                ),
                target_data AS(
                    SELECT
                        subject_expand.sems_subj_score_id,
                        COALESCE(subject_expand.student_id, graduation_plan_expand_with_student.student_id) AS student_id,
                        COALESCE(subject_expand.graduation_plan_id, graduation_plan_expand_with_student.graduation_plan_id) AS graduation_plan_id,
                        COALESCE(subject_expand.grade_year, graduation_plan_expand_with_student.grade_year) AS grade_year,
                        subject_expand.school_year AS school_year,
                        COALESCE(subject_expand.semester, graduation_plan_expand_with_student.semester) AS semester,
                        COALESCE(subject_expand.subject_name, graduation_plan_expand_with_student.subject_name) AS subject_name,
                        COALESCE(subject_expand.subject_level, graduation_plan_expand_with_student.subject_level) AS subject_level,
                        subject_expand.credit,
                        subject_expand.不計學分,
                        subject_expand.不需評分,
                        graduation_plan_expand_with_student.domain,
                        graduation_plan_expand_with_student.分組名稱,
                        graduation_plan_expand_with_student.分組修課學分數,
		                suggest_graduation_plan.subject_name AS suggest_subject_name,
		                suggest_graduation_plan.subject_level AS suggest_subject_level
                    FROM
                        subject_expand
                    FULL OUTER JOIN graduation_plan_expand_with_student
                        ON subject_expand.student_id = graduation_plan_expand_with_student.student_id
                        AND subject_expand.grade_year = graduation_plan_expand_with_student.grade_year
                        AND subject_expand.semester = graduation_plan_expand_with_student.semester
                        AND subject_expand.subject_name = graduation_plan_expand_with_student.subject_name
                        AND subject_expand.subject_level = graduation_plan_expand_with_student.subject_level
                    FULL OUTER JOIN graduation_plan_expand_with_student AS suggest_graduation_plan
                        ON subject_expand.student_id = suggest_graduation_plan.student_id
                        AND subject_expand.grade_year = suggest_graduation_plan.grade_year
                        AND subject_expand.semester = suggest_graduation_plan.semester
                        AND subject_expand.subject_name = suggest_graduation_plan.subject_name
                ),
                target_match AS (
                    --成績年級、學期、科目、級別比對到的資料
                    SELECT
                        target_data.student_id,
                        target_data.grade_year,
                        target_data.semester,
                        target_data.school_year,
                        target_data.subject_name,
                        target_data.subject_level,
                        target_data.分組名稱,
                        target_data.分組修課學分數,
                        target_data.credit
                    FROM
                        target_data
                    WHERE
                        target_data.student_id IS NOT NULL
                        AND target_data.sems_subj_score_id IS NOT NULL
                        AND target_data.分組名稱 IS NOT NULL
                    ORDER BY
                        student_id,
                        grade_year,
                        semester,
                        subject_name,
                        subject_level
                ),
                target_mismatch AS (
                    --成績年級、學期、科目、級別比對到的資料
                    SELECT
                        target_data.student_id,
                        target_data.grade_year,
                        target_data.semester,
                        target_data.school_year,
                        target_data.subject_name,
                        target_data.subject_level,
                        target_data.credit,
                        target_data.不計學分,
                        target_data.不需評分
                    FROM
                        target_data
                    WHERE
		                target_data.student_id IS NOT NULL
                        AND target_data.分組名稱 IS NULL
                    ORDER BY
                        student_id,
                        grade_year,
                        semester,
                        subject_name,
                        subject_level
                ),
                graduation_plan_mismatch_all AS(
                    SELECT
                        target_data.student_id,
                        target_data.graduation_plan_id,
                        target_data.grade_year,
                        target_data.semester,
                        target_data.domain,
                        target_data.subject_name,
                        target_data.subject_level,
                        target_data.分組名稱,
                        target_data.分組修課學分數
                    FROM
                        target_data
                    WHERE
		                target_data.student_id IS NOT NULL
                        AND target_data.sems_subj_score_id IS NULL
                ),
                graduation_plan_mismatch AS (
                    --學生的課程規劃表有，卻沒有比對到的成績年級、學期、科目、級別  
                    SELECT
                        graduation_plan_mismatch_all.*
                    FROM
                        graduation_plan_mismatch_all
                    WHERE
                        分組名稱 = ''
                ),
                graduation_plan_subject_group_mismatch AS (
                    --學生課程規劃表中，規劃該年級的課程群組中，在比對到的資料中學分總數不符合的
                    /*
                     -- 條件
                     找出課程群組中的課程，用科目名稱+級別比對實際修課或成績的學分數是否符合群組設定的學分
                     */
                    SELECT
                        target_data.student_id,
                        target_data.graduation_plan_id,
                        target_data.grade_year,
                        target_data.semester,
                        target_data.分組名稱,
                        target_data.分組修課學分數,
                        SUM(target_data.credit) AS sum_credit
                    FROM
                        target_data
                    WHERE
                        target_data.student_id IS NOT NULL
                        AND target_data.分組名稱 <> ''
                    GROUP BY
                        target_data.student_id,
                        target_data.graduation_plan_id,
                        target_data.grade_year,
                        target_data.semester,
                        target_data.分組名稱,
                        target_data.分組修課學分數
                    HAVING
                        SUM(target_data.credit) <> target_data.分組修課學分數
                )
                    -- sheet3
                    SELECT
                        target_student.student_id AS 學生系統編號,
                        target_student.student_number AS 學號,
                        target_student.class_name AS 班級,
                        target_student.dept_name AS 科別名稱,
                        target_student.seat_no AS 座號,
                        target_student.student_name AS 姓名,
                        target_student.graduation_plan_name AS 使用課程規劃表,
                        graduation_plan_subject_group_mismatch.grade_year AS 成績年級,
                        graduation_plan_subject_group_mismatch.semester AS 學期,
                        graduation_plan_subject_group_mismatch.分組名稱,
                        graduation_plan_subject_group_mismatch.分組修課學分數,
                        graduation_plan_subject_group_mismatch.sum_credit AS 成績累計學分數
                    FROM
                        graduation_plan_subject_group_mismatch
                        INNER JOIN target_student
                            ON target_student.student_id = graduation_plan_subject_group_mismatch.student_id
", condition);

                DataTable dt = qh.Select(strSQL);
                foreach (DataRow dr in dt.Rows)
                {
                    value.Add(dr);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }

        // 取得學期成績與課規以科目名稱+級別比對相同，領域、指定學年科目名稱、課程代碼、報部科目名稱不同，不同。
        public static List<DataRow> GetSemsSubjectLevelCheckGraduationPlan4(string GradeYear, string DeptID, string ClassID)
        {
            List<DataRow> value = new List<DataRow>();
            try
            {
                // 取得科別對照
                Dictionary<string, string> deptDict = GetDeptIDNameDict();


                string condition = @"
                    SELECT
                        " + GradeYear + @"::INT AS grade_year, -- NULL時為全部年級 
                        NULL :: INTEGER AS class_id,
                        NULL :: INTEGER AS dept_id
                ";


                // SELECT 3::INT AS grade_year -- NULL時為全部年級
                //	           , NULL::TEXT AS dept_name--NULL時為全部科別

                if (!string.IsNullOrEmpty(DeptID))
                {
                    condition = @"
                    SELECT                        
                        NULL :: INTEGER AS class_id,
                        " + DeptID + " AS dept_id," +
                        "" + GradeYear + " AS grade_year ";
                }


                if (!string.IsNullOrEmpty(ClassID))
                {
                    condition = @"
                    SELECT                        
                        " + ClassID + " AS class_id," +
                        "" + DeptID + " AS dept_id," +
                    "" + GradeYear + " AS grade_year ";
                }

                QueryHelper qh = new QueryHelper();
                string strSQL = string.Format(@"
                WITH row AS(
	               {0}
                ),
                target_student AS(
	                SELECT
                        student.id AS student_id,
                        graduation_plan.id AS graduation_plan_id,
                        graduation_plan.name AS graduation_plan_name,
		                class.id AS class_id,
		                class.class_name,
		                dept.id AS dept_id,
		                student.student_number,
		                student.seat_no,
		                student.name AS student_name,
                        dept.name AS dept_name
	                FROM
		                row
		                INNER JOIN class
			                       ON (
                                   class.grade_year = row.grade_year 
				                   AND ( 
                                        row.class_id is null 
                                        OR class.id = row.class_id
                                    )   
			                    )
		                INNER JOIN student
			                ON student.ref_class_id = class.id
			                AND student.status IN (1, 2)
		                INNER JOIN dept
			                ON dept.id = COALESCE(student.ref_dept_id, class.ref_dept_id)
			                AND (
				                row.dept_id IS NULL
				                 OR dept.id = row.dept_id
			                )
                        INNER JOIN graduation_plan 
                            ON graduation_plan.id = COALESCE(
                                student.ref_graduation_plan_id,
                                class.ref_graduation_plan_id
                            )
                ),
                target_student_with_sems_history AS(
                    SELECT
                        ref_student_id AS student_id,
                        graduation_plan_id,
                        grade_year :: SMALLINT,
                        semester :: SMALLINT,
                        MAX(school_year) AS school_year
                    FROM
                        sems_subj_score
                        INNER JOIN target_student
			                ON target_student.student_id = sems_subj_score.ref_student_id
                    GROUP BY
                        ref_student_id,
                        graduation_plan_id,
                        grade_year,
                        semester
                ),
                graduation_plan_expand AS(
                    SELECT
                        graduation_plan_expand.graduation_plan_id,
                        array_to_string(xpath('//Subject/@GradeYear', subject_ele), '') :: SMALLINT AS grade_year,
                        array_to_string(xpath('//Subject/@Semester', subject_ele), '') :: SMALLINT AS semester,
                        array_to_string(xpath('//Subject/@SubjectName', subject_ele), '') :: TEXT AS subject_name,
                        array_to_string(xpath('//Subject/@Level', subject_ele), '') :: TEXT AS subject_level,
                        array_to_string(xpath('//Subject/@Domain', subject_ele), '') :: TEXT AS domain,
                        array_to_string(xpath('//Subject/@分組名稱', subject_ele), '') :: TEXT AS 分組名稱,
                        (
                            '0' || array_to_string(xpath('//Subject/@分組修課學分數', subject_ele), '')
                        ) :: INTEGER AS 分組修課學分數,
                        array_to_string(xpath('//Subject/@指定學年科目名稱', subject_ele), '') :: TEXT AS 指定學年科目名稱,
				                        array_to_string(xpath('//Subject/@課程代碼', subject_ele), '') :: TEXT AS 課程代碼,
                        array_to_string(xpath('//Subject/@Entry', subject_ele), '') :: TEXT AS 分項類別,
                        array_to_string(xpath('//Subject/@RequiredBy', subject_ele), '') :: TEXT AS 校部訂,
                        array_to_string(xpath('//Subject/@Required', subject_ele), '') :: TEXT AS 必選修,
                        array_to_string(xpath('//Subject/@報部科目名稱', subject_ele), '') :: TEXT AS 報部科目名稱
                    FROM
                        (
                            SELECT
                                target_graduation.graduation_plan_id,
                                unnest(
                                    xpath(
                                        '//GraduationPlan/Subject',
                                        xmlparse(content graduation_plan.content)
                                    )
                                ) AS subject_ele
                            FROM
                                (
					                SELECT
						                DISTINCT
						                graduation_plan_id
					                FROM
						                target_student
				                ) AS target_graduation
                                INNER JOIN graduation_plan 
					                ON graduation_plan.id = target_graduation.graduation_plan_id
                        ) AS graduation_plan_expand
                ),
                graduation_plan_expand_with_student AS(
                    SELECT
                        target_student_with_sems_history.student_id,
		                graduation_plan_expand.*
                    FROM
                        graduation_plan_expand
                    INNER JOIN target_student_with_sems_history 
		                ON target_student_with_sems_history.graduation_plan_id = graduation_plan_expand.graduation_plan_id
		                AND target_student_with_sems_history.grade_year = graduation_plan_expand.grade_year
		                AND target_student_with_sems_history.semester = graduation_plan_expand.semester
    
                ),
                subject_expand AS(
                    SELECT
                        sems_subj_score_ext.sems_subj_score_id,
                        sems_subj_score_ext.student_id,
                        sems_subj_score_ext.graduation_plan_id,
                        sems_subj_score_ext.grade_year,
                        sems_subj_score_ext.school_year,
                        sems_subj_score_ext.semester,
                        array_to_string(xpath('//Subject/@科目', subj_score_ele), '') :: text AS subject_name,
                        array_to_string(xpath('//Subject/@科目級別', subj_score_ele), '') :: text AS subject_level,
                        (
                            '0' || array_to_string(xpath('//Subject/@開課學分數', subj_score_ele), '')
                        ) :: INTEGER AS credit,
                        array_to_string(xpath('//Subject/@不計學分', subj_score_ele), '')::text AS 不計學分,
                        array_to_string(xpath('//Subject/@不需評分', subj_score_ele), '')::text AS 不需評分,
	                    array_to_string(xpath('//Subject/@領域', subj_score_ele), '')::text AS 領域,
			                    array_to_string(xpath('//Subject/@修課科目代碼', subj_score_ele), '')::text AS 課程代碼,
			            array_to_string(xpath('//Subject/@指定學年科目名稱', subj_score_ele), '')::text AS 指定學年科目名稱,
                        array_to_string(xpath('//Subject/@開課分項類別', subj_score_ele), '')::text AS 分項類別,
                        array_to_string(xpath('//Subject/@修課校部訂', subj_score_ele), '')::text AS 校部訂,
                        array_to_string(xpath('//Subject/@修課必選修', subj_score_ele), '')::text AS 必選修,
                        array_to_string(xpath('//Subject/@報部科目名稱', subj_score_ele), '')::text AS 報部科目名稱
                    FROM
                        (
                            SELECT
                                sems_subj_score.id AS sems_subj_score_id,
                                sems_subj_score.ref_student_id AS student_id,
                                target_student_with_sems_history.graduation_plan_id,
                                sems_subj_score.grade_year,
                                sems_subj_score.school_year,
                                sems_subj_score.semester,
                                unnest(
                                    xpath(
                                        '//SemesterSubjectScoreInfo/Subject',
                                        xmlparse(content score_info)
                                    )
                                ) as subj_score_ele
                            FROM
                                sems_subj_score
                                INNER JOIN target_student_with_sems_history 
					                ON target_student_with_sems_history.student_id = sems_subj_score.ref_student_id
					                AND target_student_with_sems_history.grade_year = sems_subj_score.grade_year
					                AND target_student_with_sems_history.school_year = sems_subj_score.school_year
					                AND target_student_with_sems_history.semester = sems_subj_score.semester
                        ) as sems_subj_score_ext
                ),
                target_data AS(
                    SELECT
                        subject_expand.sems_subj_score_id,
                        COALESCE(subject_expand.student_id, graduation_plan_expand_with_student.student_id) AS student_id,
                        COALESCE(subject_expand.graduation_plan_id, graduation_plan_expand_with_student.graduation_plan_id) AS graduation_plan_id,
                        COALESCE(subject_expand.grade_year, graduation_plan_expand_with_student.grade_year) AS grade_year,
                        subject_expand.school_year AS school_year,
                        COALESCE(subject_expand.semester, graduation_plan_expand_with_student.semester) AS semester,
                        COALESCE(subject_expand.subject_name, graduation_plan_expand_with_student.subject_name) AS subject_name,
                        COALESCE(subject_expand.subject_level, graduation_plan_expand_with_student.subject_level) AS subject_level,
                        subject_expand.credit,
                        subject_expand.不計學分,
                        subject_expand.不需評分,
                        graduation_plan_expand_with_student.domain,
                        graduation_plan_expand_with_student.分組名稱,
                        graduation_plan_expand_with_student.分組修課學分數,
		                suggest_graduation_plan.subject_name AS suggest_subject_name,
		                suggest_graduation_plan.subject_level AS suggest_subject_level,
                        subject_expand.領域,
                        subject_expand.課程代碼,
                        subject_expand.指定學年科目名稱,
                        graduation_plan_expand_with_student.指定學年科目名稱 AS 新指定學年科目名稱,
                        graduation_plan_expand_with_student.課程代碼 AS 新課程代碼,
                        subject_expand.分項類別,
                        subject_expand.必選修,
                        subject_expand.校部訂,
                        graduation_plan_expand_with_student.分項類別 AS 新分項類別,
                        graduation_plan_expand_with_student.必選修 AS 新必選修,
                        graduation_plan_expand_with_student.校部訂 AS 新校部訂,
                        subject_expand.報部科目名稱,
                        graduation_plan_expand_with_student.報部科目名稱 AS 新報部科目名稱
                    FROM
                        subject_expand
                    FULL OUTER JOIN graduation_plan_expand_with_student
                        ON subject_expand.student_id = graduation_plan_expand_with_student.student_id
                        AND subject_expand.grade_year = graduation_plan_expand_with_student.grade_year
                        AND subject_expand.semester = graduation_plan_expand_with_student.semester
                        AND subject_expand.subject_name = graduation_plan_expand_with_student.subject_name
                        AND subject_expand.subject_level = graduation_plan_expand_with_student.subject_level
                    FULL OUTER JOIN graduation_plan_expand_with_student AS suggest_graduation_plan
                        ON subject_expand.student_id = suggest_graduation_plan.student_id
                        AND subject_expand.grade_year = suggest_graduation_plan.grade_year
                        AND subject_expand.semester = suggest_graduation_plan.semester
                        AND subject_expand.subject_name = suggest_graduation_plan.subject_name
                ),
                target_match AS (
                    --成績年級、學期、科目、級別比對到的資料
                    SELECT
                       target_data.student_id,
		                target_data.grade_year,
		                target_data.semester,
		                target_data.school_year,
		                target_data.subject_name,
		                target_data.subject_level,
		                target_data.分組名稱,
		                target_data.分組修課學分數,
		                target_data.credit,
		                target_data.領域,
		                target_data.指定學年科目名稱,
		                target_data.課程代碼,
		                target_data.domain AS 新領域,
		                target_data.新指定學年科目名稱,
		                target_data.新課程代碼,
		                target_data.報部科目名稱,
		                target_data.新報部科目名稱,
		                target_data.分項類別,
		                target_data.新分項類別,
		                target_data.校部訂,
		                target_data.新校部訂,
		                target_data.必選修,
		                target_data.新必選修
                    FROM
                        target_data
                    WHERE
                        target_data.student_id IS NOT NULL
                        AND target_data.sems_subj_score_id IS NOT NULL
                        AND target_data.分組名稱 IS NOT NULL
                    ORDER BY
                        student_id,
                        grade_year,
                        semester,
                        subject_name,
                        subject_level
                ),
                target_mismatch AS (
                    --成績年級、學期、科目、級別比對到的資料
                 	SELECT
		            target_data.student_id,
		            target_data.grade_year,
		            target_data.semester,
		            target_data.school_year,
		            target_data.subject_name,
		            target_data.subject_level,
		            target_data.credit,
		            target_data.不計學分,
		            target_data.不需評分
                    FROM
                        target_data
                    WHERE
		                target_data.student_id IS NOT NULL
                        AND target_data.分組名稱 IS NULL
                    ORDER BY
                        student_id,
                        grade_year,
                        semester,
                        subject_name,
                        subject_level
                ),
                graduation_plan_mismatch_all AS(
                    SELECT
                        target_data.student_id,
                        target_data.graduation_plan_id,
                        target_data.grade_year,
                        target_data.semester,
                        target_data.domain,
                        target_data.subject_name,
                        target_data.subject_level,
                        target_data.分組名稱,
                        target_data.分組修課學分數
                    FROM
                        target_data
                    WHERE
		                target_data.student_id IS NOT NULL
                        AND target_data.sems_subj_score_id IS NULL
                ),
                graduation_plan_mismatch AS (
                    --學生的課程規劃表有，卻沒有比對到的成績年級、學期、科目、級別  
                    SELECT
                        graduation_plan_mismatch_all.*
                    FROM
                        graduation_plan_mismatch_all
                    WHERE
                        分組名稱 = ''
                ),
                graduation_plan_subject_group_mismatch AS (
                    --學生課程規劃表中，規劃該年級的課程群組中，在比對到的資料中學分總數不符合的
                    /*
                     -- 條件
                     找出課程群組中的課程，用科目名稱+級別比對實際修課或成績的學分數是否符合群組設定的學分
                     */
                    SELECT
                        target_data.student_id,
                        target_data.graduation_plan_id,
                        target_data.grade_year,
                        target_data.semester,
                        target_data.分組名稱,
                        target_data.分組修課學分數,
                        SUM(target_data.credit) AS sum_credit
                    FROM
                        target_data
                    WHERE
                        target_data.student_id IS NOT NULL
                        AND target_data.分組名稱 <> ''
                    GROUP BY
                        target_data.student_id,
                        target_data.graduation_plan_id,
                        target_data.grade_year,
                        target_data.semester,
                        target_data.分組名稱,
                        target_data.分組修課學分數
                    HAVING
                        SUM(target_data.credit) <> target_data.分組修課學分數
                )
                 SELECT DISTINCT 
		                target_match.student_id AS 學生系統編號,
		                target_student.student_number AS 學號,
		                target_student.dept_name AS 科別名稱,
		                target_student.class_name AS 班級,
		                target_student.seat_no AS 座號,
		                target_student.student_name AS 姓名,
		                target_match.school_year AS 學年度,
		                target_match.semester AS 學期,
		                target_match.grade_year AS 成績年級,
		                target_match.subject_name AS 科目名稱,
		                target_match.subject_level AS 科目級別,
		                target_match.credit AS 學分數,
		                target_match.領域,
		                target_match.新領域,
		                target_match.指定學年科目名稱,
		                target_match.新指定學年科目名稱,
		                target_match.課程代碼,
		                target_match.新課程代碼,
		                target_match.分項類別,
		                target_match.新分項類別,
		                target_match.校部訂,
		                target_match.新校部訂,
		                target_match.必選修,
		                target_match.新必選修,                        
		                target_match.報部科目名稱,
		                target_match.新報部科目名稱
	                FROM
		                target_match 
			                INNER JOIN target_student
			                ON target_match.student_id = target_student.student_id
		                WHERE 
			                (領域 <> 新領域) 
			                OR (指定學年科目名稱 <> 新指定學年科目名稱)
			                OR (課程代碼 <> 新課程代碼) 
			                OR (分項類別 <> 新分項類別) 
			                OR (校部訂 <> 新校部訂)
			                OR (必選修 <> 新必選修)			
			                OR (報部科目名稱 <> 新報部科目名稱)
	                ORDER BY
			                班級,
			                座號,
			                學號,
			                學年度,
			                學期,
			                科目名稱 
", condition);

                // Utility.ExportText("sql4", strSQL);

                DataTable dt = qh.Select(strSQL);
                foreach (DataRow dr in dt.Rows)
                {
                    value.Add(dr);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }


        // 取得學期成績與課規以科目名稱+級別比對相同，報部科目名稱不同，不同。
        public static List<DataRow> GetSemsSubjectLevelCheckGraduationPlan5(string GradeYear, string DeptID, string ClassID)
        {
            List<DataRow> value = new List<DataRow>();
            try
            {
                // 取得科別對照
                Dictionary<string, string> deptDict = GetDeptIDNameDict();


                string condition = @"
                    SELECT
                        " + GradeYear + @"::INT AS grade_year, -- NULL時為全部年級 
                        NULL :: INTEGER AS class_id,
                        NULL :: INTEGER AS dept_id
                ";


                // SELECT 3::INT AS grade_year -- NULL時為全部年級
                //	           , NULL::TEXT AS dept_name--NULL時為全部科別

                if (!string.IsNullOrEmpty(DeptID))
                {
                    condition = @"
                    SELECT                        
                        NULL :: INTEGER AS class_id,
                        " + DeptID + " AS dept_id," +
                        "" + GradeYear + " AS grade_year ";
                }


                if (!string.IsNullOrEmpty(ClassID))
                {
                    condition = @"
                    SELECT                        
                        " + ClassID + " AS class_id," +
                        "" + DeptID + " AS dept_id," +
                    "" + GradeYear + " AS grade_year ";
                }

                QueryHelper qh = new QueryHelper();
                string strSQL = string.Format(@"
                WITH row AS(
	               {0}
                ),
                target_student AS(
	                SELECT
                        student.id AS student_id,
                        graduation_plan.id AS graduation_plan_id,
                        graduation_plan.name AS graduation_plan_name,
		                class.id AS class_id,
		                class.class_name,
		                dept.id AS dept_id,
		                student.student_number,
		                student.seat_no,
		                student.name AS student_name,
                        dept.name AS dept_name
	                FROM
		                row
		                INNER JOIN class
			                       ON (
                                   class.grade_year = row.grade_year 
				                   AND ( 
                                        row.class_id is null 
                                        OR class.id = row.class_id
                                    )   
			                    )
		                INNER JOIN student
			                ON student.ref_class_id = class.id
			                AND student.status IN (1, 2)
		                INNER JOIN dept
			                ON dept.id = COALESCE(student.ref_dept_id, class.ref_dept_id)
			                AND (
				                row.dept_id IS NULL
				                 OR dept.id = row.dept_id
			                )
                        INNER JOIN graduation_plan 
                            ON graduation_plan.id = COALESCE(
                                student.ref_graduation_plan_id,
                                class.ref_graduation_plan_id
                            )
                ),
                target_student_with_sems_history AS(
                    SELECT
                        ref_student_id AS student_id,
                        graduation_plan_id,
                        grade_year :: SMALLINT,
                        semester :: SMALLINT,
                        MAX(school_year) AS school_year
                    FROM
                        sems_subj_score
                        INNER JOIN target_student
			                ON target_student.student_id = sems_subj_score.ref_student_id
                    GROUP BY
                        ref_student_id,
                        graduation_plan_id,
                        grade_year,
                        semester
                ),
                graduation_plan_expand AS(
                    SELECT
                        graduation_plan_expand.graduation_plan_id,
                        array_to_string(xpath('//Subject/@GradeYear', subject_ele), '') :: SMALLINT AS grade_year,
                        array_to_string(xpath('//Subject/@Semester', subject_ele), '') :: SMALLINT AS semester,
                        array_to_string(xpath('//Subject/@SubjectName', subject_ele), '') :: TEXT AS subject_name,
                        array_to_string(xpath('//Subject/@Level', subject_ele), '') :: TEXT AS subject_level,
                        array_to_string(xpath('//Subject/@Domain', subject_ele), '') :: TEXT AS domain,
                        array_to_string(xpath('//Subject/@分組名稱', subject_ele), '') :: TEXT AS 分組名稱,
                        (
                            '0' || array_to_string(xpath('//Subject/@分組修課學分數', subject_ele), '')
                        ) :: INTEGER AS 分組修課學分數,
                        array_to_string(xpath('//Subject/@指定學年科目名稱', subject_ele), '') :: TEXT AS 指定學年科目名稱,
				                        array_to_string(xpath('//Subject/@課程代碼', subject_ele), '') :: TEXT AS 課程代碼,
                        array_to_string(xpath('//Subject/@Entry', subject_ele), '') :: TEXT AS 分項類別,
                        array_to_string(xpath('//Subject/@RequiredBy', subject_ele), '') :: TEXT AS 校部訂,
                        array_to_string(xpath('//Subject/@Required', subject_ele), '') :: TEXT AS 必選修,
                        array_to_string(xpath('//Subject/@報部科目名稱', subject_ele), '') :: TEXT AS 報部科目名稱
                    FROM
                        (
                            SELECT
                                target_graduation.graduation_plan_id,
                                unnest(
                                    xpath(
                                        '//GraduationPlan/Subject',
                                        xmlparse(content graduation_plan.content)
                                    )
                                ) AS subject_ele
                            FROM
                                (
					                SELECT
						                DISTINCT
						                graduation_plan_id
					                FROM
						                target_student
				                ) AS target_graduation
                                INNER JOIN graduation_plan 
					                ON graduation_plan.id = target_graduation.graduation_plan_id
                        ) AS graduation_plan_expand
                ),
                graduation_plan_expand_with_student AS(
                    SELECT
                        target_student_with_sems_history.student_id,
		                graduation_plan_expand.*
                    FROM
                        graduation_plan_expand
                    INNER JOIN target_student_with_sems_history 
		                ON target_student_with_sems_history.graduation_plan_id = graduation_plan_expand.graduation_plan_id
		                AND target_student_with_sems_history.grade_year = graduation_plan_expand.grade_year
		                AND target_student_with_sems_history.semester = graduation_plan_expand.semester
    
                ),
                subject_expand AS(
                    SELECT
                        sems_subj_score_ext.sems_subj_score_id,
                        sems_subj_score_ext.student_id,
                        sems_subj_score_ext.graduation_plan_id,
                        sems_subj_score_ext.grade_year,
                        sems_subj_score_ext.school_year,
                        sems_subj_score_ext.semester,
                        array_to_string(xpath('//Subject/@科目', subj_score_ele), '') :: text AS subject_name,
                        array_to_string(xpath('//Subject/@科目級別', subj_score_ele), '') :: text AS subject_level,
                        (
                            '0' || array_to_string(xpath('//Subject/@開課學分數', subj_score_ele), '')
                        ) :: INTEGER AS credit,
                        array_to_string(xpath('//Subject/@不計學分', subj_score_ele), '')::text AS 不計學分,
                        array_to_string(xpath('//Subject/@不需評分', subj_score_ele), '')::text AS 不需評分,
	                    array_to_string(xpath('//Subject/@領域', subj_score_ele), '')::text AS 領域,
			                    array_to_string(xpath('//Subject/@修課科目代碼', subj_score_ele), '')::text AS 課程代碼,
			            array_to_string(xpath('//Subject/@指定學年科目名稱', subj_score_ele), '')::text AS 指定學年科目名稱,
                        array_to_string(xpath('//Subject/@開課分項類別', subj_score_ele), '')::text AS 分項類別,
                        array_to_string(xpath('//Subject/@修課校部訂', subj_score_ele), '')::text AS 校部訂,
                        array_to_string(xpath('//Subject/@修課必選修', subj_score_ele), '')::text AS 必選修,
                        array_to_string(xpath('//Subject/@報部科目名稱', subj_score_ele), '')::text AS 報部科目名稱
                    FROM
                        (
                            SELECT
                                sems_subj_score.id AS sems_subj_score_id,
                                sems_subj_score.ref_student_id AS student_id,
                                target_student_with_sems_history.graduation_plan_id,
                                sems_subj_score.grade_year,
                                sems_subj_score.school_year,
                                sems_subj_score.semester,
                                unnest(
                                    xpath(
                                        '//SemesterSubjectScoreInfo/Subject',
                                        xmlparse(content score_info)
                                    )
                                ) as subj_score_ele
                            FROM
                                sems_subj_score
                                INNER JOIN target_student_with_sems_history 
					                ON target_student_with_sems_history.student_id = sems_subj_score.ref_student_id
					                AND target_student_with_sems_history.grade_year = sems_subj_score.grade_year
					                AND target_student_with_sems_history.school_year = sems_subj_score.school_year
					                AND target_student_with_sems_history.semester = sems_subj_score.semester
                        ) as sems_subj_score_ext
                ),
                target_data AS(
                    SELECT
                        subject_expand.sems_subj_score_id,
                        COALESCE(subject_expand.student_id, graduation_plan_expand_with_student.student_id) AS student_id,
                        COALESCE(subject_expand.graduation_plan_id, graduation_plan_expand_with_student.graduation_plan_id) AS graduation_plan_id,
                        COALESCE(subject_expand.grade_year, graduation_plan_expand_with_student.grade_year) AS grade_year,
                        subject_expand.school_year AS school_year,
                        COALESCE(subject_expand.semester, graduation_plan_expand_with_student.semester) AS semester,
                        COALESCE(subject_expand.subject_name, graduation_plan_expand_with_student.subject_name) AS subject_name,
                        COALESCE(subject_expand.subject_level, graduation_plan_expand_with_student.subject_level) AS subject_level,
                        subject_expand.credit,
                        subject_expand.不計學分,
                        subject_expand.不需評分,
                        graduation_plan_expand_with_student.domain,
                        graduation_plan_expand_with_student.分組名稱,
                        graduation_plan_expand_with_student.分組修課學分數,
		                suggest_graduation_plan.subject_name AS suggest_subject_name,
		                suggest_graduation_plan.subject_level AS suggest_subject_level,
                        subject_expand.領域,
                        subject_expand.課程代碼,
                        subject_expand.指定學年科目名稱,
                        graduation_plan_expand_with_student.指定學年科目名稱 AS 新指定學年科目名稱,
                        graduation_plan_expand_with_student.課程代碼 AS 新課程代碼,
                        subject_expand.分項類別,
                        subject_expand.必選修,
                        subject_expand.校部訂,
                        graduation_plan_expand_with_student.分項類別 AS 新分項類別,
                        graduation_plan_expand_with_student.必選修 AS 新必選修,
                        graduation_plan_expand_with_student.校部訂 AS 新校部訂,
                        subject_expand.報部科目名稱,
                        graduation_plan_expand_with_student.報部科目名稱 AS 新報部科目名稱
                    FROM
                        subject_expand
                    FULL OUTER JOIN graduation_plan_expand_with_student
                        ON subject_expand.student_id = graduation_plan_expand_with_student.student_id
                        AND subject_expand.grade_year = graduation_plan_expand_with_student.grade_year
                        AND subject_expand.semester = graduation_plan_expand_with_student.semester
                        AND subject_expand.subject_name = graduation_plan_expand_with_student.subject_name
                        AND subject_expand.subject_level = graduation_plan_expand_with_student.subject_level
                    FULL OUTER JOIN graduation_plan_expand_with_student AS suggest_graduation_plan
                        ON subject_expand.student_id = suggest_graduation_plan.student_id
                        AND subject_expand.grade_year = suggest_graduation_plan.grade_year
                        AND subject_expand.semester = suggest_graduation_plan.semester
                        AND subject_expand.subject_name = suggest_graduation_plan.subject_name
                ),
                target_match AS (
                    --成績年級、學期、科目、級別比對到的資料
                   SELECT
		                target_data.student_id,
		                target_data.grade_year,
		                target_data.semester,
		                target_data.school_year,
		                target_data.subject_name,
		                target_data.subject_level,
		                target_data.分組名稱,
		                target_data.分組修課學分數,
		                target_data.credit,
		                target_data.領域,
		                target_data.指定學年科目名稱,
		                target_data.課程代碼,
		                target_data.domain AS 新領域,
		                target_data.新指定學年科目名稱,
		                target_data.新課程代碼,
		                target_data.報部科目名稱,
		                target_data.新報部科目名稱,
		                target_data.分項類別,
		                target_data.新分項類別,
		                target_data.校部訂,
		                target_data.新校部訂,
		                target_data.必選修,
		                target_data.新必選修,
                        target_data.sems_subj_score_id 
                    FROM
                        target_data
                    WHERE
                        target_data.student_id IS NOT NULL
                        AND target_data.sems_subj_score_id IS NOT NULL
                        AND target_data.分組名稱 IS NOT NULL
                    ORDER BY
                        student_id,
                        grade_year,
                        semester,
                        subject_name,
                        subject_level
                ),
                target_mismatch AS (
                    --成績年級、學期、科目、級別比對到的資料
                    SELECT
                        target_data.student_id,
                        target_data.grade_year,
                        target_data.semester,
                        target_data.school_year,
                        target_data.subject_name,
                        target_data.subject_level,
                        target_data.credit,
                        target_data.不計學分,
                        target_data.不需評分
                    FROM
                        target_data
                    WHERE
		                target_data.student_id IS NOT NULL
                        AND target_data.分組名稱 IS NULL
                    ORDER BY
                        student_id,
                        grade_year,
                        semester,
                        subject_name,
                        subject_level
                ),
                graduation_plan_mismatch_all AS(
                    SELECT
                        target_data.student_id,
                        target_data.graduation_plan_id,
                        target_data.grade_year,
                        target_data.semester,
                        target_data.domain,
                        target_data.subject_name,
                        target_data.subject_level,
                        target_data.分組名稱,
                        target_data.分組修課學分數
                    FROM
                        target_data
                    WHERE
		                target_data.student_id IS NOT NULL
                        AND target_data.sems_subj_score_id IS NULL
                ),
                graduation_plan_mismatch AS (
                    --學生的課程規劃表有，卻沒有比對到的成績年級、學期、科目、級別  
                    SELECT
                        graduation_plan_mismatch_all.*
                    FROM
                        graduation_plan_mismatch_all
                    WHERE
                        分組名稱 = ''
                ),
                graduation_plan_subject_group_mismatch AS (
                    --學生課程規劃表中，規劃該年級的課程群組中，在比對到的資料中學分總數不符合的
                    /*
                     -- 條件
                     找出課程群組中的課程，用科目名稱+級別比對實際修課或成績的學分數是否符合群組設定的學分
                     */
                    SELECT
                        target_data.student_id,
                        target_data.graduation_plan_id,
                        target_data.grade_year,
                        target_data.semester,
                        target_data.分組名稱,
                        target_data.分組修課學分數,
                        SUM(target_data.credit) AS sum_credit
                    FROM
                        target_data
                    WHERE
                        target_data.student_id IS NOT NULL
                        AND target_data.分組名稱 <> ''
                    GROUP BY
                        target_data.student_id,
                        target_data.graduation_plan_id,
                        target_data.grade_year,
                        target_data.semester,
                        target_data.分組名稱,
                        target_data.分組修課學分數
                    HAVING
                        SUM(target_data.credit) <> target_data.分組修課學分數
                )
                    SELECT DISTINCT 
		                target_match.student_id AS 學生系統編號,
		                target_student.student_number AS 學號,
		                target_student.dept_name AS 科別名稱,
		                target_student.class_name AS 班級,
		                target_student.seat_no AS 座號,
		                target_student.student_name AS 姓名,
		                target_match.school_year AS 學年度,
		                target_match.semester AS 學期,
		                target_match.grade_year AS 成績年級,
		                target_match.subject_name AS 科目名稱,
		                target_match.subject_level AS 科目級別,
		                target_match.credit AS 學分數,
		                target_match.領域,
		                target_match.新領域,
		                target_match.指定學年科目名稱,
		                target_match.新指定學年科目名稱,
		                target_match.課程代碼,
		                target_match.新課程代碼,
		                target_match.分項類別,
		                target_match.新分項類別,
		                target_match.校部訂,
		                target_match.新校部訂,
		                target_match.必選修,
		                target_match.新必選修,                        
		                target_match.報部科目名稱,
		                target_match.新報部科目名稱,
                        target_match.sems_subj_score_id
	                FROM
		                target_match 
			                INNER JOIN target_student
			                ON target_match.student_id = target_student.student_id
		                WHERE 			               
                            (報部科目名稱 <> 新報部科目名稱)
	                ORDER BY
			                班級,
			                座號,
			                學號,
			                學年度,
			                學期,
			                科目名稱 
", condition);

                // Utility.ExportText("sql4", strSQL);

                DataTable dt = qh.Select(strSQL);
                foreach (DataRow dr in dt.Rows)
                {
                    value.Add(dr);
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
                        array_to_string(xpath('//Subject/@指定學年科目名稱', subject_ele), '') :: text AS 指定學年科目名稱,
                        array_to_string(xpath('//Subject/@分組名稱', subject_ele), '') :: text AS 分組名稱,
                        array_to_string(xpath('//Subject/@分組修課學分數', subject_ele), '') :: text AS 分組修課學分數,
                        array_to_string(xpath('//Subject/@設定對開', subject_ele), '') :: text AS 設定對開
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
                    gs.Group = dr["分組名稱"] + "";
                    gs.GroupCredit = dr["分組修課學分數"] + "";
                    gs.RequiredPair = dr["設定對開"] + "";

                    //// -- 使用科目名稱與級別當key
                    // string gsKey = gs.SubjectName + "_" + gs.SubjectLevel;
                    // 使用年級+學期+科目名稱當key
                    string gsKey = gs.GradeYear + "_" + gs.Semester + "_" + gs.SubjectName;
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

        // 刪除學期科目
        public static int DelSemsScoreSubject(List<StudSubjectInfo> dataList)
        {
            int value = 0;
            try
            {
                if (dataList.Count > 0)
                {
                    // 取得需要刪除學期成績系統編號
                    List<string> SemsScoreIDs = new List<string>();

                    foreach (StudSubjectInfo ss in dataList)
                    {
                        if (!SemsScoreIDs.Contains(ss.SemsSubjID))
                            SemsScoreIDs.Add(ss.SemsSubjID);
                    }
                    XElement dataXML = null;

                    string strSQL = "SELECT id,score_info FROM sems_subj_score WHERE id IN(" + string.Join(",", SemsScoreIDs.ToArray()) + ");";

                    Dictionary<string, SemsScoreInfo> SemsScoreInfoDict = new Dictionary<string, SemsScoreInfo>();

                    int count = 0;
                    QueryHelper qh = new QueryHelper();
                    DataTable dt = qh.Select(strSQL);
                    foreach (DataRow dr in dt.Rows)
                    {
                        string id = dr["id"] + "";
                        string score_info = dr["score_info"] + "";
                        try
                        {
                            if (!SemsScoreInfoDict.ContainsKey(id))
                            {
                                SemsScoreInfo sems = new SemsScoreInfo();
                                sems.id = id;
                                sems.ScoreInfo = XElement.Parse(score_info);
                                SemsScoreInfoDict.Add(id, sems);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }

                    StringBuilder sbLog = new StringBuilder();
                    sbLog.AppendLine("== 刪除學生學期科目資料 ==");

                    // 移除相對科目Element
                    foreach (StudSubjectInfo ss in dataList)
                    {
                        if (SemsScoreInfoDict.ContainsKey(ss.SemsSubjID))
                        {
                            XElement elmRoot = SemsScoreInfoDict[ss.SemsSubjID].ScoreInfo;

                            foreach (XElement elm in elmRoot.Elements("Subject"))
                            {
                                if (elm.Attribute("科目").Value == ss.SubjectName && elm.Attribute("科目級別").Value == ss.SubjectLevel)
                                {
                                    elm.Remove();
                                    sbLog.AppendLine("學年度:" + ss.SchoolYear + "，學期:" + ss.Semester + "，學號:" + ss.StudentNumber + "，班級：" + ss.ClassName + "，座號：" + ss.SeatNo + "，姓名:" + ss.Name + "，科目名稱：" + ss.SubjectName + "，級別：" + ss.SubjectLevel);
                                    count++;
                                }
                            }
                        }
                    }
                    sbLog.AppendLine("共刪除" + count + "筆。");

                    try
                    {
                        // 更新資料
                        string UpdateStrSQL = @"WITH update_data AS(";
                        int cot = 1;
                        foreach (SemsScoreInfo ssi in SemsScoreInfoDict.Values)
                        {
                            UpdateStrSQL += "SELECT " + ssi.id + " AS id,'" + ssi.ScoreInfo.ToString() + "' AS score_info";
                            if (cot < SemsScoreInfoDict.Count)
                                UpdateStrSQL += " UNION ALL ";
                            cot++;
                        }

                        UpdateStrSQL += @") ,update_sems_data AS (
                    UPDATE 
                        sems_subj_score 
                            SET score_info = update_data.score_info 
                    FROM update_data                    
                        WHERE sems_subj_score.id = update_data.id RETURNING sems_subj_score.id
                    )
                    SELECT * FROM 
                        update_sems_data";

                        DataTable dtUpdate = qh.Select(UpdateStrSQL);

                        FISCA.LogAgent.ApplicationLog.Log("成績系統.畢業預警與資料合理檢查", "刪除學生學期科目資料", sbLog.ToString());
                        value = count;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return value;
        }

        // 更新學期科目內容
        public static int UpdateSemsScoreSubjectInfo(List<StudSubjectInfo> dataList)
        {
            int value = 0;
            try
            {
                if (dataList.Count > 0)
                {
                    // 取得需要刪除學期成績系統編號
                    List<string> SemsScoreIDs = new List<string>();

                    foreach (StudSubjectInfo ss in dataList)
                    {
                        if (!SemsScoreIDs.Contains(ss.SemsSubjID))
                            SemsScoreIDs.Add(ss.SemsSubjID);
                    }
                    XElement dataXML = null;

                    string strSQL = "SELECT id,score_info FROM sems_subj_score WHERE id IN(" + string.Join(",", SemsScoreIDs.ToArray()) + ");";

                    Dictionary<string, SemsScoreInfo> SemsScoreInfoDict = new Dictionary<string, SemsScoreInfo>();

                    int count = 0;
                    QueryHelper qh = new QueryHelper();
                    DataTable dt = qh.Select(strSQL);
                    foreach (DataRow dr in dt.Rows)
                    {
                        string id = dr["id"] + "";
                        string score_info = dr["score_info"] + "";
                        try
                        {
                            if (!SemsScoreInfoDict.ContainsKey(id))
                            {
                                SemsScoreInfo sems = new SemsScoreInfo();
                                sems.id = id;
                                sems.ScoreInfo = XElement.Parse(score_info);
                                SemsScoreInfoDict.Add(id, sems);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }

                    StringBuilder sbLog = new StringBuilder();
                    sbLog.AppendLine("== 更新學生學期科目資料 ==");

                    // 更新相對科目Element
                    foreach (StudSubjectInfo ss in dataList)
                    {
                        if (SemsScoreInfoDict.ContainsKey(ss.SemsSubjID))
                        {
                            XElement elmRoot = SemsScoreInfoDict[ss.SemsSubjID].ScoreInfo;

                            foreach (XElement elm in elmRoot.Elements("Subject"))
                            {
                                if (ss.IsSubjectLevelChanged)
                                {
                                    if (ss.SubjectName == elm.Attribute("科目").Value && ss.SubjectLevel == elm.Attribute("科目級別").Value)
                                    {
                                        elm.SetAttributeValue("科目級別", ss.SubjectLevelNew);

                                        sbLog.AppendLine("學年度:" + ss.SchoolYear + "，學期:" + ss.Semester + "，學號:" + ss.StudentNumber + "，班級：" + ss.ClassName + "，座號：" + ss.SeatNo + "，姓名:" + ss.Name + "，科目名稱：" + ss.SubjectName + "，級別：由「" + ss.SubjectLevel + "」改成「" + ss.SubjectLevelNew + "」。");
                                        count++;
                                    }
                                }
                            }
                        }
                    }

                    // 回寫資料
                    sbLog.AppendLine("共更新" + count + "筆。");

                    try
                    {
                        // 更新資料
                        string UpdateStrSQL = @"WITH update_data AS(";
                        int cot = 1;
                        foreach (SemsScoreInfo ssi in SemsScoreInfoDict.Values)
                        {
                            UpdateStrSQL += "SELECT " + ssi.id + " AS id,'" + ssi.ScoreInfo.ToString() + "' AS score_info";
                            if (cot < SemsScoreInfoDict.Count)
                                UpdateStrSQL += " UNION ALL ";
                            cot++;
                        }

                        UpdateStrSQL += @") ,update_sems_data AS (
                    UPDATE 
                        sems_subj_score 
                            SET score_info = update_data.score_info 
                    FROM update_data                    
                        WHERE sems_subj_score.id = update_data.id RETURNING sems_subj_score.id
                    )
                    SELECT * FROM 
                        update_sems_data";

                        DataTable dtUpdate = qh.Select(UpdateStrSQL);

                        FISCA.LogAgent.ApplicationLog.Log("成績系統.畢業預警與資料合理檢查", "更新學生學期科目資料", sbLog.ToString());
                        value = count;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return value;
        }

        // 更新學期科目報部名稱
        public static int UpdateSemsScoreDSubjectName(List<DataRow> dataList)
        {
            int value = 0;
            try
            {
                if (dataList.Count > 0)
                {
                    // 取得需要更新學期成績系統編號
                    List<string> SemsScoreIDs = new List<string>();

                    foreach (DataRow dr in dataList)
                    {
                        string semsID = dr["sems_subj_score_id"] + "";
                        if (!SemsScoreIDs.Contains(semsID))
                            SemsScoreIDs.Add(semsID);
                    }
                    XElement dataXML = null;

                    string strSQL = "SELECT id,score_info FROM sems_subj_score WHERE id IN(" + string.Join(",", SemsScoreIDs.ToArray()) + ");";

                    Dictionary<string, SemsScoreInfo> SemsScoreInfoDict = new Dictionary<string, SemsScoreInfo>();

                    int count = 0;
                    QueryHelper qh = new QueryHelper();
                    DataTable dt = qh.Select(strSQL);
                    foreach (DataRow dr in dt.Rows)
                    {
                        string id = dr["id"] + "";
                        string score_info = dr["score_info"] + "";
                        try
                        {
                            if (!SemsScoreInfoDict.ContainsKey(id))
                            {
                                SemsScoreInfo sems = new SemsScoreInfo();
                                sems.id = id;
                                sems.ScoreInfo = XElement.Parse(score_info);
                                SemsScoreInfoDict.Add(id, sems);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }

                    StringBuilder sbLog = new StringBuilder();
                    sbLog.AppendLine("== 更新學生學期科目資料 ==");

                    // 更新相對科目Element
                    foreach (DataRow dr in dataList)
                    {
                        string semsID = dr["sems_subj_score_id"] + "";
                        if (SemsScoreInfoDict.ContainsKey(semsID))
                        {
                            XElement elmRoot = SemsScoreInfoDict[semsID].ScoreInfo;

                            foreach (XElement elm in elmRoot.Elements("Subject"))
                            {
                                // 科目名稱+級別比對
                                string SubjectName = dr["科目名稱"] + "";
                                string SubjectLevel = dr["科目級別"] + "";

                                // 報部科目名稱
                                string DSubjectNameOld = dr["報部科目名稱"] + "";
                                string DSubjectName = dr["新報部科目名稱"] + "";


                                if (SubjectName == elm.Attribute("科目").Value && SubjectLevel == elm.Attribute("科目級別").Value)
                                {
                                    elm.SetAttributeValue("報部科目名稱", DSubjectName);

                                    sbLog.AppendLine("學年度:" + dr["學年度"] + "，學期:" + dr["學期"] + "，成績年級: " + dr["成績年級"] + "，學號:" + dr["學號"] + "，班級：" + dr["班級"] + "，座號：" + dr["座號"] + "，姓名:" + dr["姓名"] + "，科目名稱：" + SubjectName + "，科目級別：" + SubjectLevel + "，報部科目名稱：由「" + DSubjectNameOld + "」改成「" + DSubjectName + "」。");
                                    count++;
                                }

                            }
                        }
                    }

                    // 回寫資料
                    sbLog.AppendLine("共更新" + count + "筆。");

                    try
                    {
                        // 更新資料
                        string UpdateStrSQL = @"WITH update_data AS(";
                        int cot = 1;
                        foreach (SemsScoreInfo ssi in SemsScoreInfoDict.Values)
                        {
                            UpdateStrSQL += "SELECT " + ssi.id + " AS id,'" + ssi.ScoreInfo.ToString() + "' AS score_info";
                            if (cot < SemsScoreInfoDict.Count)
                                UpdateStrSQL += " UNION ALL ";
                            cot++;
                        }

                        UpdateStrSQL += @") ,update_sems_data AS (
                    UPDATE 
                        sems_subj_score 
                            SET score_info = update_data.score_info 
                    FROM update_data                    
                        WHERE sems_subj_score.id = update_data.id RETURNING sems_subj_score.id
                    )
                    SELECT * FROM 
                        update_sems_data";

                        DataTable dtUpdate = qh.Select(UpdateStrSQL);

                        FISCA.LogAgent.ApplicationLog.Log("成績系統.畢業預警與資料合理檢查", "更新學生學期科目資料", sbLog.ToString());
                        value = count;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return value;
        }


        // 取得學生學期科目資料畢業使用
        public static List<StudSubjectInfo> GetSemsSubjectInfoForGraduation(string DeptID, string ClassID, string textName)
        {


            List<StudSubjectInfo> value = new List<StudSubjectInfo>();
            try
            {
                // 取得科別對照
                Dictionary<string, string> deptDict = GetDeptIDNameDict();

                string condition = " ";

                if (!string.IsNullOrEmpty(DeptID))
                    condition = " dept_id = " + DeptID + "";

                if (!string.IsNullOrEmpty(ClassID))
                    condition = " class_id = " + ClassID + "";

                if (!string.IsNullOrEmpty(textName) && !string.IsNullOrEmpty(DeptID))
                    condition += " AND student_name LIKE '" + textName + "%'";
                else
                {
                    if (string.IsNullOrEmpty(DeptID))
                        condition += " student_name LIKE '" + textName + "%'";
                }

                QueryHelper qh = new QueryHelper();
                string strSQL = string.Format(@"
                WITH student_base_source AS(
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
                    student.status IN(1,2)
            ),
            student_base AS (
                SELECT
                    *
                FROM
                    student_base_source
                WHERE
                    {0}
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
                    array_to_string(xpath('//Subject/@領域', subj_score_ele), '') :: text AS 領域,
                    array_to_string(xpath('//Subject/@不計學分', subj_score_ele), '') :: text AS 不計學分,
                    array_to_string(xpath('//Subject/@是否取得學分', subj_score_ele), '') :: text AS 是否取得學分
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
                sems_subj_score.指定學年科目名稱,
                sems_subj_score.不計學分,
                sems_subj_score.是否取得學分
            FROM
                student_base   
                INNER JOIN sems_subj_score ON student_base.student_id = sems_subj_score.ref_student_id
    
            ORDER BY
                student_base.student_number,
                sems_subj_score.科目名稱,
                sems_subj_score.科目級別,
                sems_subj_score.school_year,
                sems_subj_score.semester
", condition);
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
                    // 科別名稱
                    if (deptDict.ContainsKey(sc.DeptID))
                        sc.DeptName = deptDict[sc.DeptID];

                    sc.CoursePlanID = dr["g_plan_id"] + "";
                    sc.Pass = dr["是否取得學分"] + "";
                    sc.NotIncludedInCredit = dr["不計學分"] + "";
                    value.Add(sc);
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
