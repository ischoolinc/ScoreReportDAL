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
                int idx = 1;
                string strGearYear = "";
                char[] cp = si.credit_period.ToArray();
                foreach (char c in cp)
                {
                    string credit = c + "";

                    if (idx == 1)
                    {
                        strGearYear = "1";
                    }
                    else if (idx == 2)
                    {
                        strGearYear = "1";
                    }
                    else if (idx == 3)
                    {
                        strGearYear = "2";
                    }
                    else if (idx == 4)
                    {
                        strGearYear = "2";
                    }
                    else if (idx == 5)
                    {
                        strGearYear = "3";
                    }
                    else if (idx == 6)
                    {
                        strGearYear = "3";
                    }
                    else
                    {

                    }

                    // 學分格式0 不放入
                    if (credit == "0")
                    {
                        idx++;
                        continue;
                    }

                    string key = si.Entry + "_" + si.SubjectName.Trim() + "_" + si.RequireBy + "_" + si.Required + "_" + strGearYear;
                    //if (!SubjectInfoDict.ContainsKey(si.GetSubjectKey()))
                    //    SubjectInfoDict.Add(si.GetSubjectKey(), si);
                    if (!SubjectInfoDict.ContainsKey(key))
                        SubjectInfoDict.Add(key, si);
                }
            }

            return SubjectInfoDict;
        }

        public string GetCourseCode(string Entry, string SubjectName, string RequireBy, string Required, string GradeYear)
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
            string SubjectKey = Entry + "_" + SubjectName.Trim() + "_" + RequireBy + "_" + Required + "_" + GradeYear;

            if (SubjectInfoDict.ContainsKey(SubjectKey))
                code = SubjectInfoDict[SubjectKey].GetCourseCode();

            return code;
        }

        public void ParseSemesterHistorySchoolYear()
        {
            SemesterHistorySchoolYearDict.Clear();

            foreach (SemesterHistoryItem item in SemesterHistoryItems)
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
        public string GetScScoreSchoolYear(string SubjectName, string SubjLevel)
        {
            string value = "";

            try
            {
                string key = SubjectName.Trim() + "_" + SubjLevel;
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

