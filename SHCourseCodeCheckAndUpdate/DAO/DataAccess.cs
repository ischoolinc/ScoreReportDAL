using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using FISCA.Data;
using System.Web.Script.Serialization;
using System.Xml.Linq;
using FISCA.LogAgent;
using System.Runtime.InteropServices.ComTypes;
using Aspose.Words.Fields;
using System.Diagnostics.Eventing.Reader;

namespace SHCourseCodeCheckAndUpdate.DAO
{
    public class DataAccess
    {
        /// <summary>
        /// 傳入學年度、學期、年級，取得學生修課資料
        /// </summary>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <param name="GradeYear"></param>
        /// <returns></returns>
        public static List<StudSCAttendInfo> GetStudSCAttendBySchoolYearSems(string SchoolYear, string Semester, string GradeYear)
        {
            List<StudSCAttendInfo> value = new List<StudSCAttendInfo>();
            string strSQL = string.Format(@"
            WITH semester_data AS(
                SELECT
                    '{0}' :: INT AS school_year,
                    '{1}' :: INT AS semester,
                    '{2}' :: INT AS grade_year
            ),
            student_base AS(
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
                    END AS status
                FROM
                    student
                    LEFT JOIN class ON student.ref_class_id = class.id
                    INNER JOIN semester_data ON class.grade_year = semester_data.grade_year
                WHERE
                    student.status <> 256
                    AND class.grade_year IS NOT NULL
            ),
            course_data AS(
                SELECT
                    course.school_year AS 學年度,
                    course.semester AS 學期,
                    course.course_name AS 課程名稱,
                    course.subject AS 科目名稱,
                    course.subj_level :: TEXT AS 科目級別,
                    sc_attend.id AS sc_attend_id,
                    student_base.student_id AS student_id,
                    sc_attend.subject_code AS course_code
                FROM
                    semester_data
                    INNER JOIN course ON semester_data.school_year = course.school_year
                    AND semester_data.semester = course.semester
                    INNER JOIN sc_attend ON course.id = sc_attend.ref_course_id
                    INNER JOIN student_base ON sc_attend.ref_student_id = student_base.student_id
            ),
            gp_plan_data AS(
                SELECT
                    id,
                    moe_group_code,
                    name,
                    array_to_string(xpath('//Subject/@GradeYear', subject_ele), '') :: INT AS grade_year,
                    array_to_string(xpath('//Subject/@Semester', subject_ele), '') :: INT AS semester,
                    array_to_string(xpath('//Subject/@Entry', subject_ele), '') :: text AS 分項,
                    array_to_string(xpath('//Subject/@Domain', subject_ele), '') :: text AS 領域,
                    array_to_string(xpath('//Subject/@SubjectName', subject_ele), '') :: text AS 科目名稱,
                    array_to_string(xpath('//Subject/@Level', subject_ele), '') :: text AS 科目級別,
                    array_to_string(xpath('//Subject/@Credit', subject_ele), '') :: text AS 學分數,
                    array_to_string(xpath('//Subject/@Required', subject_ele), '') :: text AS 必選修,
                    array_to_string(xpath('//Subject/@RequiredBy', subject_ele), '') :: text AS 校部訂,
                    array_to_string(xpath('//Subject/@課程代碼', subject_ele), '') :: text AS 課程代碼
                FROM
                    (
                        SELECT
                            id,
                            name,
                            moe_group_code,
                            unnest(
                                xpath(
                                    '//GraduationPlan/Subject',
                                    xmlparse(content content)
                                )
                            ) as subject_ele
                        FROM
                            graduation_plan
                        WHERE
                            id IN(
                                SELECT
                                    DISTINCT g_plan_id
                                FROM
                                    student_base
                            )
                    ) AS graduation_plan
            ),
            student_course_code AS (
                SELECT
                    student_base.student_id AS 學生系統編號,
                    student_base.student_number AS 學號,
                    student_base.seat_no AS 座號,
                    student_base.student_name AS 姓名,
                    student_base.class_name AS 班級,
                    student_base.grade_year AS 班級年級,
                    student_base.g_plan_id AS 課程規劃編號,
                    student_base.status AS 學生狀態,
                    course_data.學年度 AS 學年度,
                    course_data.學期 AS 學期,
                    course_data.課程名稱 AS 課程名稱,
                    course_data.科目名稱 AS 科目名稱,
                    course_data.科目級別 AS 科目級別,
                    course_data.sc_attend_id AS 修課系統編號,
                    course_data.course_code AS 修課課程代碼,
                    gp_plan_data.校部訂 AS 校部訂,
                    gp_plan_data.必選修 AS 必選修,
                    gp_plan_data.學分數 AS 學分,
                    gp_plan_data.name AS 使用課程規劃表,
                    gp_plan_data.課程代碼 AS 課程規劃課程代碼
                FROM
                    semester_data
                    INNER JOIN student_base ON semester_data.grade_year = student_base.grade_year
                    INNER JOIN course_data ON student_base.student_id = course_data.student_id
                    AND course_data.學年度 = semester_data.school_year
                    AND course_data.學期 = semester_data.semester
                    INNER JOIN gp_plan_data ON student_base.g_plan_id = gp_plan_data.id
                    AND course_data.科目名稱 = gp_plan_data.科目名稱
                    AND course_data.科目級別 = gp_plan_data.科目級別
            )
            SELECT
                *
            FROM
                student_course_code
            WHERE
                修課課程代碼 <> 課程規劃課程代碼
            ORDER BY
                班級,
                座號,
                課程名稱
", SchoolYear, Semester, GradeYear);

            try
            {
                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(strSQL);
                foreach (DataRow dr in dt.Rows)
                {
                    StudSCAttendInfo sc = new StudSCAttendInfo();
                    sc.StudentID = dr["學生系統編號"] + "";
                    sc.SCAttendID = dr["修課系統編號"] + "";
                    sc.SchoolYear = dr["學年度"] + "";
                    sc.Semester = dr["學期"] + "";
                    sc.GradeYear = dr["班級年級"] + "";
                    sc.StudentNumber = dr["學號"] + "";
                    sc.ClassName = dr["班級"] + "";
                    sc.SeatNo = dr["座號"] + "";
                    sc.Name = dr["姓名"] + "";
                    sc.CourseName = dr["課程名稱"] + "";
                    sc.SubjectName = dr["科目名稱"] + "";
                    sc.SubjectLevel = dr["科目級別"] + "";
                    sc.RequiredBy = dr["校部訂"] + "";
                    sc.Required = dr["必選修"] + "";
                    sc.Credit = dr["學分"] + "";
                    sc.SC_CourseCode = dr["修課課程代碼"] + "";
                    sc.GP_CourseCode = dr["課程規劃課程代碼"] + "";
                    sc.GPName = dr["使用課程規劃表"] + "";
                    sc.status = dr["學生狀態"] + "";
                    value.Add(sc);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }

        /// <summary>
        /// 更新修課課程代碼
        /// </summary>
        /// <param name="dataList"></param>
        /// <returns></returns>
        public static List<int> UpdateSCAttendCourseCode(List<StudSCAttendInfo> dataList)
        {
            List<int> value = new List<int>();

            List<UpdateSCAttendInfo> UpdateSCAttendInfoList = new List<UpdateSCAttendInfo>();

            foreach (StudSCAttendInfo data in dataList)
            {
                UpdateSCAttendInfo ui = new UpdateSCAttendInfo();
                ui.id = int.Parse(data.SCAttendID);
                ui.course_code = data.GP_CourseCode;
                UpdateSCAttendInfoList.Add(ui);
            }

            JavaScriptSerializer js = new JavaScriptSerializer();
            string JSonStr = js.Serialize(UpdateSCAttendInfoList);

            string UpdateSQL = @"
            WITH j_source AS (
                SELECT
                    so.json ->> 'id' AS id,
                    so.json ->> 'course_code' AS course_code
                FROM
                    (
                        select
                            json_array_elements(
                                '" + JSonStr + @"' :: JSON
                            ) AS json
                    ) AS so
            ),
            update_data AS(
                UPDATE
                    sc_attend 
                SET
                    subject_code = j_source.course_code
                FROM
                    j_source
                WHERE
                    sc_attend.id = j_source.id :: int RETURNING sc_attend.id
            )
            SELECT
                *
            FROM
                update_data;
            ";

            try
            {
                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(UpdateSQL);
                foreach (DataRow dr in dt.Rows)
                {
                    value.Add(int.Parse(dr["id"] + ""));
                }

                // 寫入log
                StringBuilder sbLog = new StringBuilder();
                sbLog.AppendLine("== 學生修課課程代碼修改 ==");
                foreach (StudSCAttendInfo ss in dataList)
                {
                    sbLog.AppendLine("學年度:" + ss.SchoolYear + "，學期:" + ss.Semester + "，課程名稱:" + ss.CourseName + "，修課系統編號：" + ss.SCAttendID + "，科目名稱：" + ss.SubjectName + "，科目級別：" + ss.SubjectLevel + "，課程代碼由「" + ss.SC_CourseCode + "」改成「" + ss.GP_CourseCode + "」。");
                }
                sbLog.AppendLine("共更新" + value.Count + "筆。");

                FISCA.LogAgent.ApplicationLog.Log("課程代碼-資料檢查", "修改修課課程代碼", sbLog.ToString());

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }

        /// <summary>
        /// 傳入學年度、學期、年級，取得學生科目成績
        /// </summary>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <param name="GradeYear"></param>
        /// <returns></returns>
        public static List<StudSubjectScoreInfo> GetSubjectScoreBySchoolYearSems(string SchoolYear, string Semester, string GradeYear)
        {
            List<StudSubjectScoreInfo> value = new List<StudSubjectScoreInfo>();
            string strSQL = string.Format(@"
            WITH semester_data AS(
                SELECT
                    '{0}' :: INT AS school_year,
                    '{1}' :: INT AS semester,
                    '{2}' :: INT AS grade_year
            ),
            student_base AS(
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
                    END AS status
                FROM
                    student
                    LEFT JOIN class ON student.ref_class_id = class.id                   
                WHERE
                    student.status <> 256
                    AND class.grade_year IS NOT NULL
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
                    array_to_string(xpath('//Subject/@不計學分', subj_score_ele), '') :: text AS 不計學分,
                    array_to_string(xpath('//Subject/@不需評分', subj_score_ele), '') :: text AS 不需評分,
                    array_to_string(xpath('//Subject/@開課學分數', subj_score_ele), '') :: text AS 開課學分數,
                    array_to_string(xpath('//Subject/@修課科目代碼', subj_score_ele), '') :: text AS 課程代碼
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
                            INNER JOIN semester_data ON sems_subj_score.school_year = semester_data.school_year
                            AND sems_subj_score.semester = semester_data.semester
                            AND sems_subj_score.grade_year = semester_data.grade_year
                            INNER JOIN student_base ON sems_subj_score.ref_student_id = student_base.student_id
                    ) as sems_subj_score_ext
            ),
            gp_plan_data AS(
                SELECT
                    id,
                    moe_group_code,
                    name,
                    array_to_string(xpath('//Subject/@GradeYear', subject_ele), '') :: INT AS grade_year,
                    array_to_string(xpath('//Subject/@Semester', subject_ele), '') :: INT AS semester,
                    array_to_string(xpath('//Subject/@Entry', subject_ele), '') :: text AS 分項,
                    array_to_string(xpath('//Subject/@Domain', subject_ele), '') :: text AS 領域,
                    array_to_string(xpath('//Subject/@SubjectName', subject_ele), '') :: text AS 科目名稱,
                    array_to_string(xpath('//Subject/@Level', subject_ele), '') :: text AS 科目級別,
                    array_to_string(xpath('//Subject/@Credit', subject_ele), '') :: text AS 學分數,
                    array_to_string(xpath('//Subject/@Required', subject_ele), '') :: text AS 必選修,
                    array_to_string(xpath('//Subject/@RequiredBy', subject_ele), '') :: text AS 校部訂,
                    array_to_string(xpath('//Subject/@課程代碼', subject_ele), '') :: text AS 課程代碼
                FROM
                    (
                        SELECT
                            id,
                            name,
                            moe_group_code,
                            unnest(
                                xpath(
                                    '//GraduationPlan/Subject',
                                    xmlparse(content content)
                                )
                            ) as subject_ele
                        FROM
                            graduation_plan
                        WHERE
                            id IN(
                                SELECT
                                    DISTINCT g_plan_id
                                FROM
                                    student_base
                            )
                    ) AS graduation_plan
            ),
            student_course_code AS (
                SELECT
                    student_base.student_id AS 學生系統編號,
                    student_base.student_number AS 學號,
                    student_base.seat_no AS 座號,
                    student_base.student_name AS 姓名,
                    student_base.class_name AS 班級,
                    sems_subj_score.grade_year AS 成績年級,
                    student_base.g_plan_id AS 課程規劃編號,
                    student_base.status AS 學生狀態,
                    sems_subj_score.id AS 學期成績系統編號,
                    sems_subj_score.school_year AS 學年度,
                    sems_subj_score.semester AS 學期,
                    sems_subj_score.科目名稱 AS 科目名稱,
                    sems_subj_score.科目級別 AS 科目級別,
                    (
                        CASE
                            sems_subj_score.校部訂
                            WHEN '部訂' THEN '部定'
                            ELSE sems_subj_score.校部訂
                        END
                    ) AS 校部訂,
                    sems_subj_score.必選修 AS 必選修,
                    sems_subj_score.開課學分數 AS 學分數,
                    sems_subj_score.課程代碼 AS 學期科目課程代碼,
                    gp_plan_data.name AS 使用課程規劃表,
                    gp_plan_data.課程代碼 AS 課程規劃課程代碼
                FROM
                 semester_data 
		            INNER JOIN sems_subj_score 
                    ON sems_subj_score.school_year = semester_data.school_year
                    AND sems_subj_score.semester = semester_data.semester 
		            AND sems_subj_score.grade_year = semester_data.grade_year 
		            INNER JOIN student_base ON sems_subj_score.ref_student_id = student_base.student_id 
                    INNER JOIN gp_plan_data ON student_base.g_plan_id = gp_plan_data.id
                    AND sems_subj_score.科目名稱 = gp_plan_data.科目名稱
                    AND sems_subj_score.科目級別 = gp_plan_data.科目級別
            )
            SELECT
                *
            FROM
                student_course_code
            WHERE
                學期科目課程代碼 <> 課程規劃課程代碼
            ORDER BY
                班級,
                座號,
                科目名稱
", SchoolYear, Semester, GradeYear);

            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(strSQL);
            foreach (DataRow dr in dt.Rows)
            {
                StudSubjectScoreInfo ss = new StudSubjectScoreInfo();
                // 學生系統編號
                ss.StudentID = dr["學生系統編號"] + "";
                // 學號
                ss.StudentNumber = dr["學號"] + "";
                // 座號
                ss.SeatNo = dr["座號"] + "";
                // 姓名
                ss.Name = dr["姓名"] + "";
                // 班級
                ss.ClassName = dr["班級"] + "";
                // 成績年級
                ss.GradeYear = dr["成績年級"] + "";
                // 課程規劃編號
                // 學生狀態
                ss.Status = dr["學生狀態"] + "";
                // 學期成績系統編號
                ss.SemsSubjID = dr["學期成績系統編號"] + "";
                // 學年度
                ss.SchoolYear = dr["學年度"] + "";
                // 學期
                ss.Semester = dr["學期"] + "";
                // 科目名稱
                ss.SubjectName = dr["科目名稱"] + "";
                // 科目級別
                ss.SubjectLevel = dr["科目級別"] + "";
                // 校部訂
                ss.RequiredBy = dr["校部訂"] + "";
                // 必選修
                ss.Required = dr["必選修"] + "";
                // 學分數
                ss.Credit = dr["學分數"] + "";
                // 學期科目課程代碼
                ss.SS_CourseCode = dr["學期科目課程代碼"] + "";
                // 使用課程規劃表
                ss.GPName = dr["使用課程規劃表"] + "";
                // 課程規劃課程代碼
                ss.GP_CourseCode = dr["課程規劃課程代碼"] + "";
                // 學生狀態
                ss.Status = dr["學生狀態"] + "";

                value.Add(ss);
            }

            return value;
        }


        public static int UpdateSubjectScoreCourseCode(Dictionary<string, List<StudSubjectScoreInfo>> dataDict)
        {
            int updateSubjectCount = 0;

            if (dataDict.Count > 0)
            {
                // 需要處理學期成績
                Dictionary<string, SemsScoreInfo> SemsScoreInfoDict = new Dictionary<string, SemsScoreInfo>();


                // 透過學期成績系統編號，取得需要更新資料
                string strSQL = @"
                SELECT 
                    id
                    ,score_info
                FROM 
                sems_subj_score 
                WHERE id IN(" + string.Join(",", dataDict.Keys.ToArray()) + ");";

                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(strSQL);

                foreach (DataRow dr in dt.Rows)
                {
                    string id = dr["id"] + "";
                    string scoreInfoStr = dr["score_info"] + "";

                    if (!SemsScoreInfoDict.ContainsKey(id))
                    {
                        SemsScoreInfo ss = new SemsScoreInfo();
                        ss.id = id;
                        try
                        {
                            ss.ScoreInfo = XElement.Parse(scoreInfoStr);
                            SemsScoreInfoDict.Add(id, ss);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }

                    }
                }

                // 比對來源與更新資料
                foreach (string semsId in dataDict.Keys)
                {
                    if (SemsScoreInfoDict.ContainsKey(semsId))
                    {
                        foreach (StudSubjectScoreInfo ss in dataDict[semsId])
                        {
                            foreach (XElement elm in SemsScoreInfoDict[semsId].ScoreInfo.Elements("Subject"))
                            {
                                // 使用科目+科目級別 當key來更新
                                if (ss.SubjectName == elm.Attribute("科目").Value && ss.SubjectLevel == elm.Attribute("科目級別").Value)
                                {
                                    elm.SetAttributeValue("修課科目代碼", ss.GP_CourseCode);
                                    updateSubjectCount++;
                                }
                            }
                        }
                    }
                }

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

                try
                {
                    DataTable dtUpdate = qh.Select(UpdateStrSQL);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                // 記錄 log
                StringBuilder sbLog = new StringBuilder();
                sbLog.AppendLine("== 調整學生學期課程代碼 ==");
                foreach (string key in dataDict.Keys)
                {
                    foreach (StudSubjectScoreInfo ss in dataDict[key])
                    {
                        sbLog.AppendLine("學年度:" + ss.SchoolYear + "，學期:" + ss.Semester + "，學號:" + ss.StudentNumber + "，班級：" + ss.ClassName + "，座號：" + ss.SeatNo + "，姓名:" + ss.Name + "，科目名稱：" + ss.SubjectName + "，科目級別：" + ss.SubjectLevel + "，課程代碼由「" + ss.SS_CourseCode + "」改成「" + ss.GP_CourseCode + "」。");
                    }
                }
                sbLog.AppendLine("共更新" + updateSubjectCount + "筆。");

                FISCA.LogAgent.ApplicationLog.Log("課程代碼-資料檢查", "修改學期課程代碼", sbLog.ToString());
            }

            return updateSubjectCount;
        }
    }
}