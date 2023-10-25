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
using Aspose.Cells;
using System.IO;

namespace SHGraduationWarning.UIForm
{
    public partial class frmCheckStudGPlanSemsScoreCourseCode : BaseForm
    {

        // 分批傳送用 List 內是 StudentID
        Dictionary<int, List<string>> StudentIDDs;
        List<DataRow> DataList;
        BackgroundWorker bgWorker;
        Workbook wb;
        Dictionary<string, int> _ColIdxDict;


        public frmCheckStudGPlanSemsScoreCourseCode(List<string> StudentIDs)
        {
            InitializeComponent();
            StudentIDDs = new Dictionary<int, List<string>>();
            DataList = new List<DataRow>();
            bgWorker = new BackgroundWorker();
            wb = new Workbook();
            _ColIdxDict = new Dictionary<string, int>();
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

            if (e.Error == null)
            {
                try
                {
                    Workbook wb1 = e.Result as Workbook;
                    if (wb1 != null)
                    {
                        Utility.ExportXls2003("學生學期科目與課規課程代碼差異", wb1);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            else
            {
                MsgBox.Show(e.Error.Message);
            }


            btnRun.Enabled = true;
        }

        private void BgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            int rp = 5;
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
                        student_number,
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
                        array_to_string(xpath('//Subject/@開課分項類別', subj_score_ele), '') :: text AS 分項類別,
                        array_to_string(xpath('//Subject/@修課校部訂', subj_score_ele), '') :: text AS 校部訂,
                        array_to_string(xpath('//Subject/@修課必選修', subj_score_ele), '') :: text AS 必選修,
                        array_to_string(xpath('//Subject/@開課學分數', subj_score_ele), '') :: text AS 學分數,
                        array_to_string(xpath('//Subject/@是否取得學分', subj_score_ele), '') :: text AS 取得學分,
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
                    student_data.grade_year AS 學生年級,
                    student_data.class_name AS 班級,
                    student_data.seat_no AS 座號,
                    student_data.student_name AS 姓名,
                    student_data.student_number AS 學號,
                    sems_subj_score.id AS 成績系統編號,
                    sems_subj_score.school_year AS 學年度,
                    sems_subj_score.semester AS 學期,
                    sems_subj_score.grade_year AS 成績年級,                    
                    sems_subj_score.科目,
                    sems_subj_score.科目級別,
                    sems_subj_score.分項類別,
                    sems_subj_score.校部訂,
                    sems_subj_score.必選修,
                    sems_subj_score.學分數,
                    sems_subj_score.取得學分,
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
                    sems_subj_score.課程代碼 <> gplan_data.課程代碼 
                ORDER BY
                    student_data.grade_year,
                    class_name,
                    seat_no
                ", string.Join(",", ids.ToArray()));

                DataTable dt = qh.Select(qry);
                foreach (DataRow dr in dt.Rows)
                    DataList.Add(dr);


                rp += 10;
                if (rp > 95)
                    rp = 95;
            }

            wb = new Workbook();
            wb = new Workbook(new MemoryStream(Properties.Resources.學生學期科目與課規課程代碼差異樣板));
            Worksheet wst = wb.Worksheets["學生學期科目與課規課程代碼"];
            int rowIdx = 1;
            _ColIdxDict.Clear();
            // 讀取欄位與索引            
            for (int co = 0; co <= wst.Cells.MaxDataColumn; co++)
            {
                _ColIdxDict.Add(wst.Cells[0, co].StringValue, co);
            }

            try
            {
                foreach (DataRow dr in DataList)
                {
                    // 學生系統編號
                    wst.Cells[rowIdx, _ColIdxDict["學生系統編號"]].PutValue(dr["學生系統編號"] + "");
                    
                    // 學生年級
                    wst.Cells[rowIdx, _ColIdxDict["學生年級"]].PutValue(dr["學生年級"] + "");

                    // 班級
                    wst.Cells[rowIdx, _ColIdxDict["班級"]].PutValue(dr["班級"] + "");

                    // 座號
                    wst.Cells[rowIdx, _ColIdxDict["座號"]].PutValue(dr["座號"] + "");

                    // 姓名
                    wst.Cells[rowIdx, _ColIdxDict["姓名"]].PutValue(dr["姓名"] + "");

                    // 成績年級
                    wst.Cells[rowIdx, _ColIdxDict["成績年級"]].PutValue(dr["成績年級"] + "");

                    // 成績學年度
                    wst.Cells[rowIdx, _ColIdxDict["學年度"]].PutValue(dr["學年度"] + "");

                    // 成績學期
                    wst.Cells[rowIdx, _ColIdxDict["學期"]].PutValue(dr["學期"] + "");

                    // 科目
                    wst.Cells[rowIdx, _ColIdxDict["科目"]].PutValue(dr["科目"] + "");

                    // 科目級別
                    wst.Cells[rowIdx, _ColIdxDict["科目級別"]].PutValue(dr["科目級別"] + "");

                    // 學期科目課程代碼                    
                    wst.Cells[rowIdx, _ColIdxDict["學期科目課程代碼"]].PutValue(dr["學期科目課程代碼"] + "");

                    // 課程代碼                    
                    wst.Cells[rowIdx, _ColIdxDict["課程代碼"]].PutValue(dr["課程代碼"] + "");

                    // 課程規劃表名稱                    
                    wst.Cells[rowIdx, _ColIdxDict["課程規劃表名稱"]].PutValue(dr["課程規劃表名稱"] + "");

                    // 分項類別                    
                    wst.Cells[rowIdx, _ColIdxDict["分項類別"]].PutValue(dr["分項類別"] + "");

                    
                    string str = dr["校部訂"] + "";
                    if (str == "部訂")
                        str = "部定";

                    // 校部訂   1                 
                    wst.Cells[rowIdx, _ColIdxDict["校部訂"]].PutValue(str);

                    // 必選修                    
                    wst.Cells[rowIdx, _ColIdxDict["必選修"]].PutValue(dr["必選修"] + "");

                    // 學分數                    
                    wst.Cells[rowIdx, _ColIdxDict["學分數"]].PutValue(dr["學分數"] + "");

                    // 取得學分                    
                    wst.Cells[rowIdx, _ColIdxDict["取得學分"]].PutValue(dr["取得學分"] + "");

                    // 取得學分                    
                    wst.Cells[rowIdx, _ColIdxDict["取得學分"]].PutValue(dr["取得學分"] + "");

                    // 學號                    
                    wst.Cells[rowIdx, _ColIdxDict["學號"]].PutValue(dr["學號"] + "");

                    rowIdx++;
                }

                wst.AutoFitColumns();
                e.Result = wb;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
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
