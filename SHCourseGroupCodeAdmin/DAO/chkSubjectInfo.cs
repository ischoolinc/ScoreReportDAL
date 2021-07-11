using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseGroupCodeAdmin.DAO
{
    /// <summary>
    /// 檢查課程規劃表與大表差異使用
    /// </summary>
    public class chkSubjectInfo
    {
        public string SubjectName { get; set; }
        public string CourseCode { get; set; }
        public string Entry { get; set; }
        public string Domain { get; set; }
        public string isRequired { get; set; }
        public string RequiredBy { get; set; }
        public string OpenStatus { get; set; }
        public string Credit { get; set; }
        public string credit_period { get; set; }
        public string course_attr { get; set; }
        public string ProcessStatus { get; set; }
        public List<string> DiffStatusList = new List<string>();
    }
}
