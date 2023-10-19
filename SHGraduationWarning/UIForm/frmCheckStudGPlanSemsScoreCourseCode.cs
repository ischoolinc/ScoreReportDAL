using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using FISCA.Data;

namespace SHGraduationWarning.UIForm
{
    public partial class frmCheckStudGPlanSemsScoreCourseCode : BaseForm
    {

        // 分批傳送用 List 內是 StudentID
        Dictionary<int, List<string>> StudentIDDs;
        List<DataRow> DataList;
        BackgroundWorker bgWorker;

        public frmCheckStudGPlanSemsScoreCourseCode(List<string> StudentIDs)
        {
            InitializeComponent();
            StudentIDDs = new Dictionary<int, List<string>>();
            DataList = new List<DataRow>();
            bgWorker = new BackgroundWorker();
            bgWorker.DoWork += BgWorker_DoWork;
            bgWorker.RunWorkerCompleted += BgWorker_RunWorkerCompleted;
            bgWorker.ProgressChanged += BgWorker_ProgressChanged;
            bgWorker.WorkerReportsProgress = true;

            // 將學生30人分批
            int idx = 0, i = 0; ;
            foreach (string id in StudentIDs)
            {
                if (i >= 30)
                {
                    idx++;
                    i = 0;
                }

                if (!StudentIDDs.ContainsKey(idx))
                    StudentIDDs.Add(idx, new List<string>());

                StudentIDDs[idx].Add(id);
                i++;
            }

        }

        private void BgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("取得資料比對中...", e.ProgressPercentage);
        }

        private void BgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
            btnRun.Enabled = true;
        }

        private void BgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            int rp = 1;
            bgWorker.ReportProgress(rp);

            DataList.Clear();
            QueryHelper qh = new QueryHelper();

            foreach (int idx in StudentIDDs.Keys)
            {
                List<string> ids = StudentIDDs[idx];
                string qry = string.Format(@"
                WITH student_base AS(
                    SELECT
                        student.id AS student_id,
                        student.name AS student_name,
                        student.student_number,
                        class.class_name,
                        student.seat_no,
                        class.grade_year,
                        COALESCE(student.gdc_code, class.gdc_code) AS s_gdc_code
                    FROM
                        student
                        LEFT JOIN class ON student.ref_class_id = class.id
                    WHERE
                        student.id IN(
                            {0}
                        )
                ),
                student_sems_history AS(
                    SELECT
                        history.id AS student_id,
                        array_to_string(xpath('//History/@SchoolYear', history_xml), '') :: INT AS school_year,
                        array_to_string(xpath('//History/@Semester', history_xml), '') :: INT AS semester,
                        array_to_string(xpath('//History/@GradeYear', history_xml), '') :: INT AS sh_grade_year,
                        array_to_string(xpath('//History/@GDCCode', history_xml), '') :: text AS sh_gdc_code
                    FROM
                        (
                            SELECT
                                student.id,
                                unnest(
                                    xpath(
                                        '//root/History',
                                        xmlparse(content '<root>' || sems_history || '</root>')
                                    )
                                ) AS history_xml
                            FROM
                                student
                                INNER JOIN student_base ON student.id = student_base.student_id
                        ) AS history
                ),
                student_data AS(
                    SELECT
                        student_sems_history.school_year,
                        student_sems_history.semester,
                        student_base.student_id,
                        COALESCE(
                            student_sems_history.sh_grade_year,
                            student_base.grade_year
                        ) AS grade_year,
                        seat_no,
                        student_name,
                        class_name,
                        COALESCE(sh_gdc_code, s_gdc_code) AS gdc_code
                    FROM
                        student_base
                        INNER JOIN student_sems_history ON student_base.student_id = student_sems_history.student_id
                ),
                sems_subj_score AS (
                    SELECT
                        sems_subj_score_ext.id,
                        sems_subj_score_ext.school_year,
                        sems_subj_score_ext.semester,
                        sems_subj_score_ext.grade_year,
                        sems_subj_score_ext.ref_student_id,
                        array_to_string(xpath('//Subject/@科目', subj_score_ele), '') :: text AS 科目,
                        array_to_string(xpath('//Subject/@科目級別', subj_score_ele), '') :: text AS 科目級別,
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
                                ) AS subj_score_ele
                            FROM
                                sems_subj_score
                                INNER JOIN student_data ON sems_subj_score.school_year = student_data.school_year
                                AND sems_subj_score.semester = student_data.semester
                                AND sems_subj_score.grade_year = student_data.grade_year
                                AND sems_subj_score.ref_student_id = student_data.student_id
                        ) AS sems_subj_score_ext
                ),
                gplan_data AS(
                    SELECT
                        id,
                        moe_group_code,
                        name,
                        array_to_string(xpath('//Subject/@SubjectName', subject_ele), '') :: text AS 科目,
                        array_to_string(xpath('//Subject/@Level', subject_ele), '') :: text AS 科目級別,
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
                                ) AS subject_ele
                            FROM
                                graduation_plan
                        ) AS graduation_plan
                    WHERE
                        moe_group_code IN(
                            SELECT
                                DISTINCT gdc_code
                            FROM
                                student_data
                        )
                )
                SELECT
                    student_data.student_id AS 學生系統編號,
                    student_data.school_year AS 學年度,
                    student_data.semester AS 學期,
                    student_data.grade_year AS 學生年級,
                    student_data.class_name AS 班級,
                    student_data.seat_no AS 座號,
                    student_data.student_name AS 姓名,
                    sems_subj_score.id AS 成績系統編號,
                    sems_subj_score.grade_year AS 成績年級,
                    sems_subj_score.semester AS 成績學期,
                    sems_subj_score.科目,
                    sems_subj_score.科目級別,
                    sems_subj_score.課程代碼 AS 學期科目課程代碼,
                    gplan_data.課程代碼 AS 課程代碼,
                    gplan_data.name AS 課程規劃表名稱,
                    gplan_data.moe_group_code AS 課程規劃表群科班代碼
                FROM
                    student_data
                    INNER JOIN sems_subj_score ON student_data.student_id = sems_subj_score.ref_student_id
                    AND student_data.school_year = sems_subj_score.school_year
                    AND student_data.semester = sems_subj_score.semester
                    AND student_data.grade_year = sems_subj_score.grade_year
                    LEFT JOIN gplan_data ON student_data.gdc_code = gplan_data.moe_group_code
                    AND sems_subj_score.科目 = gplan_data.科目
                    AND sems_subj_score.科目級別 = gplan_data.科目級別
                WHERE
                    sems_subj_score.課程代碼 = ''
                ORDER BY
                    student_data.grade_year,
                    class_name,
                    seat_no
                ", string.Join(",", ids.ToArray()));

                DataTable dt = qh.Select(qry);
                foreach(DataRow dr in dt.Rows)
                    DataList.Add(dr);


                rp += 10;
                if (rp > 95)
                    rp = 95;
            }

            bgWorker.ReportProgress(100);
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            btnRun.Enabled = false;
            bgWorker.RunWorkerAsync();

        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmCheckStudGPlanSemsScoreCourseCode_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
        }
    }
}
