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
                foreach(string sid in value.Keys)
                {
                    foreach(StudentCourseCodeInfo scc in value[sid])
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

            DataTable dt = qh.Select(query);
            foreach (DataRow dr in dt.Rows)
            {
                string groupCode = dr["group_code"] + "";
                if (!value.ContainsKey(groupCode))
                    value.Add(groupCode, new List<SubjectInfo>());

                SubjectInfo si = new SubjectInfo(dr["subject_name"] + "", dr["entry_year"] + "", dr["require_by"] + "", dr["is_required"] + "");
                si.SetCode(dr["group_code"] + "", dr["course_code"] + "");

                value[groupCode].Add(si);

            }

            return value;
        }



        public List<SubjectInfo> ParseCourseCodeData(List<DataRow> dataRows)
        {
            List<SubjectInfo> value = new List<SubjectInfo>();

            return value;
        }

    }
}