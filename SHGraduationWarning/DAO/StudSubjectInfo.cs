using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHGraduationWarning.DAO
{
    public class StudSubjectInfo
    {
        public string StudentID { get; set; } // 學生系統編號
        public string SemsSubjID { get; set; } // 學期成績系統編號
        public string SchoolYear { get; set; } // 學年度
        public string Semester { get; set; } // 學期

        // 班級年級
        public string ClassGradeYear { get; set; }

        public string GradeYear { get; set; } // 成績年級
        public string StudentNumber { get; set; } // 學號
        public string ClassName { get; set; } // 班級
        public string SeatNo { get; set; } // 座號
        public string Name { get; set; } // 姓名
        public string SubjectName { get; set; } // 科目名稱
        public string SubjectLevel { get; set; } // 科目級別
        public string SubjectLevelNew { get; set; } // 新科目級別

        public string Domain { get; set; } // 領域

        public string Entry { get; set; } // 分項類別

        public string RequiredBy { get; set; } // 校部定
        public string Required { get; set; } // 必選修
        public string Credit { get; set; } // 學分      
        public string GPName { get; set; } // 使用課程規畫表

        // 指定學年科目名稱
        public string SchoolYearSubjectName { get; set; }

        public string Status { get; set; } // 學生狀態

        public List<string> ErrorMsgList = new List<string>();

        // 科別
        public string DeptName { get; set; }
        // 科別系統編號
        public string DeptID { get; set; }
        // 班級系統編號
        public string ClassID { get; set; }

        // 課程規劃系統編號
        public string CoursePlanID { get; set; }

        public string GPDomain { get; set; } // 課規領域
        public string GPEntry { get; set; } // 課規分項
        public string GPCourseCode { get; set; } // 課規課程代碼

        public string GPRequiredBy { get; set; } // 課規校部定
        public string GPRequired { get; set; } // 課規必選修
        public string GPCredit { get; set; } // 課規學分
        public string GPSYSubjectName { get; set; } // 課規指定學年度科目名稱

        public bool IsSubjectLevelChanged = false; // 科目級別是否有變更 
    }
}
