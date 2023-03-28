using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseCodeCheckAndUpdate.DAO
{
    public class StudSubjectScoreInfo
    {
        public string StudentID { get; set; } // 學生系統編號
        public string SemsSubjID { get; set; } // 學期成績系統編號
        public string SchoolYear { get; set; } // 學年度
        public string Semester { get; set; } // 學期
        public string GradeYear { get; set; } // 成績年級
        public string StudentNumber { get; set; } // 學號
        public string ClassName { get; set; } // 班級
        public string SeatNo { get; set; } // 座號
        public string Name { get; set; } // 姓名
        public string SubjectName { get; set; } // 科目名稱
        public string SubjectLevel { get; set; } // 科目級別
        public string RequiredBy { get; set; } // 校部定
        public string Required { get; set; } // 必選修
        public string Credit { get; set; } // 學分
        public string SS_CourseCode { get; set; } // 學期科目課程代碼
        public string GP_CourseCode { get; set; } // 課程規劃課程代碼
        public string GPName { get; set; } // 使用課程規畫表

        public string Status { get; set; } // 學生狀態
    }
}
