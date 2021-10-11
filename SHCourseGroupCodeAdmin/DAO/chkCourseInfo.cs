using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseGroupCodeAdmin.DAO
{
    public class chkCourseInfo
    {
        public string CourseID { get; set; }
        public string CourseName { get; set; }
        public string SubjectName { get; set; }

        public string required_by { get; set; }
        public string required { get; set; }

        public string credit { get; set; }
    }
}
