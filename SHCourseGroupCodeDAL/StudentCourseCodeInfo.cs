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

        public string StudentID { get; set; }

        public string CourseGroupCode { get; set; }

        List<SubjectInfo> SubjectInfoList = new List<SubjectInfo>();

        public void AddSubjectInfo(SubjectInfo subj)
        {
            SubjectInfoList.Add(subj);
        }

        public Dictionary<string, SubjectInfo> GetSubjectInfoDict()
        {
            Dictionary<string, SubjectInfo> value = new Dictionary<string, SubjectInfo>();

            foreach (SubjectInfo si in SubjectInfoList)
            {
                if (!value.ContainsKey(si.GetSubjectKey()))
                    value.Add(si.GetSubjectKey(), si);
            }

            return value;
        }

        public void ParseCourseCode(Dictionary<string, string> SourceDict)
        {
            foreach (SubjectInfo si in SubjectInfoList)
            {
                if (SourceDict.ContainsKey(si.GetSubjectKey()))
                {
                    // 比對資料後填入
                    si.SetCode(CourseGroupCode, SourceDict[si.GetSubjectKey()]);
                }

            }
        }
    }
}

