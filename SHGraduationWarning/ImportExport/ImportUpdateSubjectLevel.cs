using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Campus.DocumentValidator;
using Campus.Import2014;
using DevComponents.DotNetBar;
using FISCA.Data;
using K12.Data;
using SHGraduationWarning.DAO;

namespace SHGraduationWarning.ImportExport
{
    public class ImportUpdateSubjectLevel : ImportWizard
    {
        // 設定檔
        private ImportOption _Option;

        public override ImportAction GetSupportActions()
        {
            //更新
            return ImportAction.Update;
        }

        public ImportUpdateSubjectLevel()
        {
            this.IsSplit = false;
        }

        public override string GetValidateRule()
        {
            // 資料驗證規格 XML 
            return Properties.Resources.ImportUpdateSubjectLevelVal;
        }

        public override string Import(List<Campus.DocumentValidator.IRowStream> Rows)
        {
            // 取得匯入資料Excel 檔案內資料
            List<StudUpdateSubjectLevelInfo> StudUpdateSubjectLevelList = new List<StudUpdateSubjectLevelInfo>();
            foreach (IRowStream ir in Rows)
            {
                // 必要欄位
                if (ir.Contains("學生系統編號") && ir.Contains("學年度") && ir.Contains("學期") && ir.Contains("成績年級") && ir.Contains("科目名稱") && ir.Contains("新科目名稱") && ir.Contains("科目級別") && ir.Contains("新科目級別"))
                {
                    string StudentID = ir.GetValue("學生系統編號");

                    StudUpdateSubjectLevelInfo ss = new StudUpdateSubjectLevelInfo();
                    ss.StudentID = StudentID;
                    ss.StudentNumber = ir.GetValue("學號");
                    ss.StudentName = ir.GetValue("姓名");
                    ss.ClassName = ir.GetValue("班級");
                    ss.SeatNo = ir.GetValue("座號");
                    ss.SchoolYear = ir.GetValue("學年度");
                    ss.Semester = ir.GetValue("學期");
                    ss.GradeYear = ir.GetValue("成績年級");
                    ss.SubjectName = ir.GetValue("科目名稱");
                    ss.SubjectNameNew = ir.GetValue("新科目名稱");
                    ss.SubjectLevel = ir.GetValue("科目級別");
                    ss.SubjectLevelNew = ir.GetValue("新科目級別");

                    StudUpdateSubjectLevelList.Add(ss);

                }
            }
            StringBuilder sbSelSemsScore = new StringBuilder();
            int cot = 1;
            foreach (StudUpdateSubjectLevelInfo ss in StudUpdateSubjectLevelList)
            {
                sbSelSemsScore.AppendLine("SELECT ");
                sbSelSemsScore.AppendLine(ss.StudentID + " AS student_id, ");
                sbSelSemsScore.AppendLine(ss.SchoolYear + " AS school_year, ");
                sbSelSemsScore.AppendLine(ss.Semester + " AS semester, ");
                sbSelSemsScore.AppendLine(ss.GradeYear + " AS grade_year ");
                if (cot < StudUpdateSubjectLevelList.Count)
                    sbSelSemsScore.AppendLine("UNION ALL ");

                cot++;
            }


            // 透過 Excel 取得資訊取得相對學生學期成績資料
            string strStudSemsScore = string.Format(@"
            WITH row AS(
	            {0}
            ),
            stud_sems_score AS (
	            SELECT
		            DISTINCT id,
		            sems_subj_score.school_year,
		            sems_subj_score.semester,
		            sems_subj_score.grade_year,
                    sems_subj_score.ref_student_id AS student_id,
		            score_info
	            FROM
		            sems_subj_score
		            INNER JOIN row 
			            ON sems_subj_score.ref_student_id = row.student_id
			            AND sems_subj_score.school_year = row.school_year
			            AND sems_subj_score.semester = row.semester
			            AND sems_subj_score.grade_year = row.grade_year
            )
            SELECT
	            *
            FROM
	            stud_sems_score
", sbSelSemsScore.ToString());

            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(strStudSemsScore);

            Dictionary<string, SemsScoreInfo> SemsScoreDict = new Dictionary<string, SemsScoreInfo>();

            // 使用學生系統編號+學年度+學期+成績年級，比對資料
            foreach (DataRow dr in dt.Rows)
            {
                SemsScoreInfo ss = new SemsScoreInfo();
                ss.StudentID = dr["student_id"] + "";
                ss.id = dr["id"] + "";
                ss.SchoolYear = dr["school_year"] + "";
                ss.Semester = dr["semester"] + "";
                ss.GradeYear = dr["grade_year"] + "";
                try
                {
                    ss.ScoreInfo = XElement.Parse(dr["score_info"] + "");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                string key = ss.StudentID + "_" + ss.SchoolYear + "_" + ss.Semester + "_" + ss.GradeYear;
                if (!SemsScoreDict.ContainsKey(key))
                    SemsScoreDict.Add(key, ss);
            }

            StringBuilder sbLog = new StringBuilder();
            sbLog.AppendLine("== 更新學生學期科目資料 ==");
            int count = 0;
            // 使用學生系統編號+學年度+學期+成績年級，比對資料整理
            foreach (StudUpdateSubjectLevelInfo StudS in StudUpdateSubjectLevelList)
            {
                string key = StudS.StudentID + "_" + StudS.SchoolYear + "_" + StudS.Semester + "_" + StudS.GradeYear;

                if (SemsScoreDict.ContainsKey(key))
                {
                    // 比對更新科目與科目級別
                    if (SemsScoreDict[key].ScoreInfo != null)
                        foreach (XElement elm in SemsScoreDict[key].ScoreInfo.Elements("Subject"))
                        {
                            if (elm.Attribute("科目").Value == StudS.SubjectName && elm.Attribute("科目級別").Value == StudS.SubjectLevel)
                            {
                                // 新科目名稱
                                elm.SetAttributeValue("科目", StudS.SubjectNameNew);

                                // 新科目級別
                                elm.SetAttributeValue("科目級別", StudS.SubjectLevelNew);

                                sbLog.AppendLine(@"
                                學生系統編號:" + StudS.StudentID + "" +
                                "，學年度:" + StudS.SchoolYear + "" +
                                "，學期:" + StudS.Semester + "" +
                                "，學號:" + StudS.StudentNumber + "" +
                                "，班級：" + StudS.ClassName + "" +
                                "，座號：" + StudS.SeatNo + "" +
                                "，姓名:" + StudS.StudentName + "" +
                                "，科目名稱：由「" + StudS.SubjectName + "」改成「" + StudS.SubjectNameNew + "」" +
                                "，級別：由「" + StudS.SubjectLevel + "」改成「" + StudS.SubjectLevelNew + "」。");

                                count++;
                            }
                        }
                }
            }
            sbLog.AppendLine("共更新" + count + "筆。");

            // 資料回寫             
            try
            {
                if (count > 0)
                {
                    // 更新資料
                    string UpdateStrSQL = @"WITH update_data AS(";
                    cot = 1;
                    foreach (SemsScoreInfo ssi in SemsScoreDict.Values)
                    {
                        UpdateStrSQL += "SELECT " + ssi.id + " AS id,'" + ssi.ScoreInfo.ToString() + "' AS score_info";
                        if (cot < SemsScoreDict.Count)
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

                    //FISCA.LogAgent.ApplicationLog.Log("成績系統.匯入更新學期科目級別", "更新學生學期科目資料", sbLog.ToString());
                    return sbLog.ToString();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return ex.Message;
            }
            return "";
        }

        public override void Prepare(ImportOption Option)
        {
            // 畫面初始化後需要載入

            _Option = Option;

        }
    }
}
