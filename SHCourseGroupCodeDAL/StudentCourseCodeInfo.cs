using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseGroupCodeDAL
{
    public class StudentCourseCodeInfo
    {
        public string SchoolYear { get; set; }
        public string Semester { get; set; }

        public string GradeYear { get; set; }

        public string StudentID { get; set; }

        public string CourseGroupCode { get; set; }

        List<SubjectInfo> SubjectInfoList = new List<SubjectInfo>();

        Dictionary<string, SubjectInfo> SubjectInfoDict = new Dictionary<string, SubjectInfo>();

        public void AddSubjectInfo(SubjectInfo subj)
        {
            SubjectInfoList.Add(subj);
        }

        public void AddSubjectInfoList(List<SubjectInfo> subjList)
        {
            SubjectInfoList.AddRange(subjList);
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

        public string GetCourseCode(string SubjectName,string RequireBy,string Required)
        {
            //  // 科目名稱+校部定+必選修
            string code = "";
            string SubjectKey = SubjectName + "_" + RequireBy + "_" + Required;

            if (SubjectInfoDict.ContainsKey(SubjectKey))
                code = SubjectInfoDict[SubjectKey].GetCourseCode();

            return code;
        }
    }
}

