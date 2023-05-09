using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHGraduationWarning.DAO
{
    // 更新學生學期科目級別
    public class StudUpdateSubjectLevelInfo
    {
        public string StudentID { get; set; }
        public string StudentNumber { get; set; }
        public string StudentName { get; set; }
        public string ClassName { get; set; }
        public string SeatNo { get; set; }
        public string SchoolYear { get; set; }
        public string Semester { get; set; }
        public string GradeYear { get; set; }
        public string SubjectName { get; set; }

        public string SubjectNameNew { get; set; }
        public string SubjectLevel { get; set; }
        public string SubjectLevelNew { get; set; }

    }
}
