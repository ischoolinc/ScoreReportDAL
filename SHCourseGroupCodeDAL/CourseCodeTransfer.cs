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

        public Dictionary<string, List<StudentCourseCodeInfo>> GetStundetCourseCodeDict(List<string> StudentIDList)
        {
            // StudentID,
            Dictionary<string, List<StudentCourseCodeInfo>> value = new Dictionary<string, List<StudentCourseCodeInfo>>();


            if (StudentIDList.Count > 0)
            {
                // 取得學生學期對照學年度、學期、群組代碼
                string queryStudS = "SELECT id,  sems_history  FROM student WHERE id IN(" + string.Join(",", StudentIDList.ToArray()) + ");";


                // < History ClassName = "" CourseGroupCode = "123456" DeptName = "" GradeYear = "1" SchoolDayCount = "" SchoolYear = "97" SeatNo = "" Semester = "1" Teacher = "" />

                QueryHelper qh = new QueryHelper();
                DataTable dtS = qh.Select(queryStudS);
                foreach (DataRow dr in dtS.Rows)
                {
                    string sid = dr["id"] + "";

                    if (!value.ContainsKey(sid))
                        value.Add(sid, new List<StudentCourseCodeInfo>());

                    // 解析學期對照取得課程代碼索引
                    try
                    {
                        XElement semsHXML = XElement.Parse("<root>" + dr["sems_history"] + "</root>");
                        if (semsHXML != null)
                        {
                            foreach (XElement elm in semsHXML.Elements("History"))
                            {
                                StudentCourseCodeInfo studInfo = new StudentCourseCodeInfo();
                                studInfo.StudentID = sid;

                                // 初始化
                                studInfo.SchoolYear = studInfo.Semester = studInfo.GradeYear = studInfo.CourseGroupCode = "";

                                if (elm.Attribute("SchoolYear") != null)
                                    studInfo.SchoolYear = elm.Attribute("SchoolYear").Value;

                                if (elm.Attribute("Semester") != null)
                                    studInfo.Semester = elm.Attribute("Semester").Value;

                                if (elm.Attribute("GradeYear") != null)
                                    studInfo.GradeYear = elm.Attribute("GradeYear").Value;

                                if (elm.Attribute("CourseGroupCode") != null)
                                    studInfo.CourseGroupCode = elm.Attribute("CourseGroupCode").Value;

                                value[sid].Add(studInfo);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }


                // 收集學生有課程群組代碼，找到相關科目對照
                List<string> CourseGroupCodeList = new List<string>();
                foreach (string key in value.Keys)
                {
                    foreach (StudentCourseCodeInfo stud in value[key])
                    {
                        if (string.IsNullOrEmpty(stud.CourseGroupCode))
                            continue;

                        if (!CourseGroupCodeList.Contains(stud.CourseGroupCode))
                            CourseGroupCodeList.Add(stud.CourseGroupCode);
                    }
                }

                // 群組所屬科目代碼
                Dictionary<string, List<SubjectInfo>> SubjectCodeDict = GetSubjectCodeDictByGroupCode(CourseGroupCodeList);

                // 解析並填入學生科目
                foreach (string sid in value.Keys)
                {
                    foreach (StudentCourseCodeInfo scc in value[sid])
                    {
                        if (SubjectCodeDict.ContainsKey(scc.CourseGroupCode))
                        {
                            scc.AddSubjectInfoList(SubjectCodeDict[scc.CourseGroupCode]);
                        }
                    }
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
            string query = "SELECT group_code,course_code,subject_name,credit_period,entry_year,require_by,is_required FROM $moe.subjectcode WHERE group_code IN('" + string.Join("''", code.ToArray()) + "')";
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


                    SubjectInfo si = new SubjectInfo(dr["subject_name"] + "", dr["entry_year"] + "", dr["require_by"] + "", dr["is_required"] + "");
                    si.SetCode(dr["group_code"] + "", dr["course_code"] + "");

                    value[groupCode].Add(si);

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


            return value;
        }



        public List<SubjectInfo> ParseCourseCodeData(List<DataRow> dataRows)
        {
            List<SubjectInfo> value = new List<SubjectInfo>();

            return value;
        }

        /// <summary>
        /// 透過學生系統編號，目前系統預設學年度學期來自學生課程規劃表對照後的課程代碼
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public Dictionary<string, StudentCourseCodeInfo> GetStudentCurrentCourseCodeDict(List<string> StudentIDList)
        {
            Dictionary<string, StudentCourseCodeInfo> value = new Dictionary<string, StudentCourseCodeInfo>();

            // 取得學生使用課程規劃表，並使用系統預設學年度、學期與學生目前班級年級，StudentCourseCodeInfo
            string query = "" +
                "SELECT " +
                "student.id AS student_id" +
                ",COALESCE(class.grade_year,0) AS grade_year" +
                ",COALESCE(student.ref_graduation_plan_id,class.ref_graduation_plan_id) as graduation_id " +
                " FROM student " +
                " LEFT JOIN class ON student.ref_class_id = class.id " +
                " WHERE student.id IN(" + string.Join(",", StudentIDList.ToArray()) + ")" +
                " ORDER BY student.student_number;";

            try
            {
                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(query);

                // 取得學生課程規劃
                foreach (DataRow dr in dt.Rows)
                {
                    StudentCourseCodeInfo scc = new StudentCourseCodeInfo();
                    scc.SchoolYear = K12.Data.School.DefaultSchoolYear;
                    scc.Semester = K12.Data.School.DefaultSemester;
                    scc.GradeYear = dr["grade_year"] + "";
                    scc.StudentID = dr["student_id"] + "";
                    scc.GraduationPlanID = dr["graduation_id"] + "";

                    if (!value.ContainsKey(scc.StudentID))
                        value.Add(scc.StudentID, new StudentCourseCodeInfo());
                }

                // 收集有使用課程規劃ID
                List<string> gpIDlist = new List<string>();

                foreach (StudentCourseCodeInfo scc in value.Values)
                {
                    if (string.IsNullOrWhiteSpace(scc.GraduationPlanID))
                        continue;

                    if (!gpIDlist.Contains(scc.GraduationPlanID))
                        gpIDlist.Add(scc.GraduationPlanID);
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }

        /// <summary>
        ///  透過課程規畫表ID 取得符合課程代碼使用課程規劃
        /// </summary>
        /// <param name="GraduationPlanIDs"></param>
        /// <returns></returns>
        public Dictionary<string, GraduationPlanInfo_code> GetGraduationPlanInfoDictByGPIDList(List<string> GraduationPlanIDs)
        {
            Dictionary<string, GraduationPlanInfo_code> value = new Dictionary<string, GraduationPlanInfo_code>();

            try
            {
                // 取得課程規劃表上群組代碼
                string qry = "SELECT " +
                    "id" +
                    ",name" +
                    ",moe_group_code" +
                    ",moe_group_code_1 " +
                    " FROM graduation_plan " +
                    " WHERE id IN(" + string.Join(",", GraduationPlanIDs.ToArray()) + ")";

                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(qry);
                List<string> groupCodeList = new List<string>();

                foreach (DataRow dr in dt.Rows)
                {
                    GraduationPlanInfo_code gpi = new GraduationPlanInfo_code();
                    gpi.ID = dr["id"] + "";
                    gpi.Name = dr["name"] + "";
                    gpi.MoeGroupCode = gpi.MoeGroupCode_1 = "";
                    if (dr["moe_group_code"] != null)
                        gpi.MoeGroupCode = dr["moe_group_code"] + "";

                    if (dr["moe_group_code_1"] != null)
                        gpi.MoeGroupCode_1 = dr["moe_group_code_1"] + "";

                    if (!value.ContainsKey(gpi.ID))
                    {
                        value.Add(gpi.ID, gpi);
                        if (!string.IsNullOrWhiteSpace(gpi.MoeGroupCode))
                        {
                            if (!groupCodeList.Contains(gpi.MoeGroupCode))
                                groupCodeList.Add(gpi.MoeGroupCode);
                        }

                        if (!string.IsNullOrWhiteSpace(gpi.MoeGroupCode_1))
                        {
                            if (!groupCodeList.Contains(gpi.MoeGroupCode_1))
                                groupCodeList.Add(gpi.MoeGroupCode_1);
                        }
                    }

                }

                Dictionary<string, List<SubjectInfo>> SubjectCodeDict = GetSubjectCodeDictByGroupCode(groupCodeList);





            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }
    }
}