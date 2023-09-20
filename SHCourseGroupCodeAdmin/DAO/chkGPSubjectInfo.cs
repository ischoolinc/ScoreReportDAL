using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseGroupCodeAdmin.DAO
{
    /// <summary>
    /// 課程規劃表科目資訊(比對使用)
    /// </summary>
    public class chkGPSubjectInfo
    {
        public string SubjectName { get; set; }
        public string CourseCode { get; set; }
        public string Entry { get; set; }
        public string Domain { get; set; }
        public string isRequired { get; set; }
        public string RequiredBy { get; set; }    
        public string Credit { get; set; }
        public string credit_period { get; set; }
        //public string course_attr { get; set; }
        public string open_type { get; set; }

        // 報部科目名稱
        public string OfficialSubjectName { get; set; }
    }
}
