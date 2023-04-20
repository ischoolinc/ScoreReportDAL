using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Data;
using System.Data;
using System.Xml.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Net.Http.Headers;
using DevComponents.Schedule.Model;

namespace SHSemsSubjectCheckEdit.DAO
{
    public class DataAccess
    {
        // 取得學期成績學年度
        public static List<string> GetSemsScoreSchoolYear()
        {
            List<string> value = new List<string>();
            try
            {
                string strSQL = @"
            SELECT
                DISTINCT school_year
            FROM
                sems_subj_score
            WHERE
                school_year > 106
            ORDER BY
                school_year DESC
            ";

                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(strSQL);
                foreach (DataRow dr in dt.Rows)
                    value.Add(dr["school_year"] + "");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }

        // 依學年度、年級，取得學期科目相關資料
        public static List<StudSubjectInfo> GetSemsSubjectBySchoolYear(string SchoolYear, int GradeYear)
        {
            List<StudSubjectInfo> value = new List<StudSubjectInfo>();

            try
            {
                if (string.IsNullOrWhiteSpace(SchoolYear))
                    return value;

                QueryHelper qh = new QueryHelper();

                for (int ss = 1; ss <= 2; ss++)
                {
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
                            student.status <> 256 AND student_number = '910029' 
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
                            array_to_string(xpath('//Subject/@開課學分數', subj_score_ele), '') :: text AS 開課學分數,
                            array_to_string(xpath('//Subject/@指定學年科目名稱', subj_score_ele), '') :: text AS 指定學年科目名稱
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
                    student_sems_subject AS (
                        SELECT
                            student_base.student_id AS 學生系統編號,
                            student_base.student_number AS 學號,
                            student_base.seat_no AS 座號,
                            student_base.student_name AS 姓名,
                            student_base.class_name AS 班級,
                            student_base.status AS 學生狀態,
                            student_base.g_plan_id AS 課程規劃表編號,
                            sems_subj_score.grade_year AS 成績年級,
                            sems_subj_score.id AS 學期成績系統編號,
                            sems_subj_score.school_year AS 學年度,
                            sems_subj_score.semester AS 學期,
                            sems_subj_score.科目名稱 AS 科目名稱,
                            sems_subj_score.科目級別 AS 科目級別,
                            sems_subj_score.校部訂,
                            sems_subj_score.必選修 AS 必選修,
                            sems_subj_score.開課學分數 AS 學分數,
                            sems_subj_score.指定學年科目名稱 AS 指定學年科目名稱
                        FROM
                            semester_data
                            INNER JOIN sems_subj_score ON sems_subj_score.school_year = semester_data.school_year
                            AND sems_subj_score.semester = semester_data.semester
                            AND sems_subj_score.grade_year = semester_data.grade_year
                            INNER JOIN student_base ON sems_subj_score.ref_student_id = student_base.student_id
                    )
                    SELECT
                        *
                    FROM
                        student_sems_subject WHERE 指定學年科目名稱 = '' 
                    ORDER BY
                        班級,
                        座號,
                        科目名稱
", SchoolYear, ss, GradeYear);
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
                        sc.RequiredBy = dr["校部訂"] + "";
                        sc.Required = dr["必選修"] + "";
                        sc.Credit = dr["學分數"] + "";
                        sc.SYSubjectName = dr["指定學年科目名稱"] + "";
                        sc.status = dr["學生狀態"] + "";
                        sc.SemsSubjID = dr["學期成績系統編號"] + "";
                        sc.GPID = dr["課程規劃表編號"] + "";
                        value.Add(sc);
                    }
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
                        array_to_string(xpath('//Subject/@課程代碼', subject_ele), '') :: text AS 課程代碼        
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

                    // 使用科目名稱與級別當key
                    string gsKey = gs.SubjectName + "_" + gs.SubjectLevel;
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

        // 刪除學生學期科目資料
        public static void DelSemsScoreSubject(List<StudSubjectInfo> dataList)
        {
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

                    // 移除相對科目Element
                    foreach(StudSubjectInfo ss in dataList)
                    {
                        if (SemsScoreInfoDict.ContainsKey(ss.SemsSubjID))
                        {
                            XElement elmRoot = SemsScoreInfoDict[ss.SemsSubjID].ScoreInfo;
                            
                            foreach(XElement elm in elmRoot.Elements("Subject"))
                            {
                                if (elm.Attribute("SubjectName").Value ==ss.SubjectName && elm.Attribute("Level").Value == ss.SubjectLevel )
                                {
                                    elm.Remove();
                                    count++;
                                }
                            }
                        }
                    }

                    // 回寫資料

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

