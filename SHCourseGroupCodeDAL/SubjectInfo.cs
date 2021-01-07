using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseGroupCodeDAL
{
    public class SubjectInfo
    {
        public SubjectInfo(string subjectName, string entryYear, string requireBy, string required)
        {
            SubjectName = subjectName;
            EntryYear = entryYear;
            RequireBy = requireBy;
            Required = required;

            // 科目名稱+校部定+必選修
            SubjectKey = SubjectName + "_" + RequireBy + "_" + Required;
        }


        /// <summary>
        /// 比對用
        /// </summary>
        private string SubjectKey = "";
        /// <summary>
        /// 科目名稱
        /// </summary>
        private string SubjectName = "";

        /// <summary>
        /// 入學年
        /// </summary>
        private string EntryYear = "";

        /// <summary>
        /// 校部定
        /// </summary>
        private string RequireBy = "";

        /// <summary>
        /// 必選修
        /// </summary>
        private string Required = "";

        /// <summary>
        /// 群組代碼
        /// </summary>
        private string GroupCode = "";


        /// <summary>
        /// 課程代碼
        /// </summary>
        private string CourseCode = "";

        public string GetSubjectKey()
        {
            return SubjectKey;
        }

        public string GetCourseCode()
        {
            return CourseCode;
        }


        public string GetSubjectName()
        {
            return SubjectName;
        }

        public string GetRequireBy() { return RequireBy; }

        public string GetRequired() { return Required; }


        public void SetCode(string groupCode, string courseCode)
        {
            GroupCode = groupCode;
            CourseCode = courseCode;
        }
    }
}
