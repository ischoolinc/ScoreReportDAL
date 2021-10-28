using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseGroupCodeAdmin.DAO
{
    public class SubjectCourseInfo
    {
        public string SchoolYear { get; set; }

        public string Semester { get; set; }

        public string SubjectName { get; set; }

        public string OpenSemester { get; set; }

        /// <summary>
        /// 班級名稱,ID 對照
        /// </summary>
        public Dictionary<string, string> ClassNameDict = new Dictionary<string, string>();

        /// <summary>
        /// 班級ID,班級所屬學生ID對照
        /// </summary>
        public Dictionary<string, List<string>> ClassStudentIDDict = new Dictionary<string, List<string>>();

        public int CourseCount { get; set; }
    }
}
