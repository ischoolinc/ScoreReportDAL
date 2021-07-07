using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseGroupCodeAdmin.DAO
{
    public class rptSubjectInfo
    {
        public string MOECode { get; set; }
        public string GroupName { get; set; }
        public string Domain { get; set; }
        public string ScoreType { get; set; }
        public string SubjectName { get; set; }
        public string isRequired { get; set; }
        public string RequiredBy { get; set; }
        public string CourseCode { get; set; }

        public string credit_period { get; set; }

        public string[] credit = new string[6];

    }
}
