using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseGroupCodeAdmin.DAO
{
    public class rptGPlanCourseChkInfo
    {
        public string SchoolYear { get; set; }
        public string Semester { get; set; }
        public string CourseID { get; set; }
        public string CourseName { get; set; }
        public string RefClassID { get; set; }
        public string ClassName { get; set; }
        public string SubjectName { get; set; }
        public string SubjectLevel { get; set; }
        public string RequiredBy { get; set; }
        public string isRequired { get; set; }
        public string Credit { get; set; }
        public string CourseCode { get; set; }
        public string credit_period { get; set; }
        public string gdc_code { get; set; }

        public List<string> ErrorMsgList = new List<string>();
    }
}
