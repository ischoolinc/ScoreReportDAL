using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Data;
using System.Data;

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
                            sems_subj_score.grade_year AS 成績年級,
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
                        student_sems_subject 
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

    }
}
