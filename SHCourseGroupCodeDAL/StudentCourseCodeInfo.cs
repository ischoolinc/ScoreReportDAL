using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using K12.Data;

namespace SHCourseGroupCodeDAL
{
    public class StudentCourseCodeInfo
    {
        public string SchoolYear { get; set; }
        public string Semester { get; set; }


        public string StudentID { get; set; }

        public string CourseGroupCode { get; set; }

        List<SubjectInfo> SubjectInfoList = new List<SubjectInfo>();

        // 學生學期對照
        public List<SemesterHistoryItem> SemesterHistoryItems = new List<SemesterHistoryItem>();

        // 應修學年度
        Dictionary<string, int> SemesterHistorySchoolYearDict = new Dictionary<string, int>();

        Dictionary<string, SubjectInfo> SubjectInfoDict = new Dictionary<string, SubjectInfo>();

        public Dictionary<string, string> ScSubjectSemesterDict = new Dictionary<string, string>();


        public void AddSubjectInfo(SubjectInfo subj)
        {
            SubjectInfoList.Add(subj);
        }

        public void AddSubjectInfoList(List<SubjectInfo> subjList)
        {
            SubjectInfoList.AddRange(subjList);
            GetSubjectInfoDict();
        }

        public Dictionary<string, SubjectInfo> GetSubjectInfoDict()
        {
            SubjectInfoDict.Clear();

            foreach (SubjectInfo si in SubjectInfoList)
            {
                if (!SubjectInfoDict.ContainsKey(si.GetSubjectKey()))
                    SubjectInfoDict.Add(si.GetSubjectKey(), si);
            }

            return SubjectInfoDict;
        }

        public string GetCourseCode(string SubjectName, string RequireBy, string Required)
        {
            // 當沒有資料時
            if (SubjectInfoDict.Count == 0)
            {
                GetSubjectInfoDict();
            }

            // 部定轉換
            if (RequireBy == "部訂")
                RequireBy = "部定";

            // 科目名稱+校部定+必選修
            string code = "";
            string SubjectKey = SubjectName + "_" + RequireBy + "_" + Required;

            if (SubjectInfoDict.ContainsKey(SubjectKey))
                code = SubjectInfoDict[SubjectKey].GetCourseCode();

            return code;
        }

        public void ParseSemesterHistorySchoolYear()
        {
            SemesterHistorySchoolYearDict.Clear();

            foreach(SemesterHistoryItem item in SemesterHistoryItems)
            {
                string str = "";
                if (item.GradeYear == 1 && item.Semester == 1)                
                    str = "1";

                if (item.GradeYear == 1 && item.Semester == 2)
                    str = "2";

                if (item.GradeYear == 2 && item.Semester == 1)
                    str = "3";
                if (item.GradeYear == 2 && item.Semester == 2)
                    str = "4";
                if (item.GradeYear == 3 && item.Semester == 1)
                    str = "5";
                if (item.GradeYear == 3 && item.Semester == 2)
                    str = "6";

                if (!SemesterHistorySchoolYearDict.ContainsKey(str))
                    SemesterHistorySchoolYearDict.Add(str, 0);

                if (item.SchoolYear > SemesterHistorySchoolYearDict[str])
                    SemesterHistorySchoolYearDict[str] = item.SchoolYear;

            }
        }


        
        /// <summary>
        /// 取得補修學年度
        /// </summary>
        /// <param name="SubjectName"></param>
        /// <param name="SubjLevel"></param>
        /// <returns></returns>
        public string GetScScoreSchoolYear(string SubjectName,  string SubjLevel)
        {
            string value = "";

            try
            {
                string key = SubjectName + "_" + SubjLevel;
                if (ScSubjectSemesterDict.ContainsKey(key))
                {
                    if (SemesterHistorySchoolYearDict.ContainsKey(ScSubjectSemesterDict[key]))
                        value = SemesterHistorySchoolYearDict[ScSubjectSemesterDict[key]].ToString();
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

