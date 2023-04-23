using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHSemsSubjectCheckEdit.DAO
{
    public class StudSubjectInfo
    {
        public string StudentID { get; set; } // 學生系統編號
        public string SemsSubjID { get; set; } // 學期成績系統編號
        public string SchoolYear { get; set; } // 學年度
        public string Semester { get; set; } // 學期
        public string GradeYear { get; set; } // 成績年級
        public string ClassGradeYear { get; set; } // 班級年級
        public string StudentNumber { get; set; } // 學號
        public string ClassName { get; set; } // 班級
        public string SeatNo { get; set; } // 座號
        public string Name { get; set; } // 姓名
        public string SubjectName { get; set; } // 科目名稱
        public string SubjectLevel { get; set; } // 科目級別
        public string RequiredBy { get; set; } // 校部定
        public string Required { get; set; } // 必選修
        public string Credit { get; set; } // 學分
        public string SYSubjectName { get; set; } // 指定學年度科目名稱

        public string status { get; set; } // 學生狀態

        public string GPID { get; set; } // 課程規劃表系統編號

        public string GPRequiredBy { get; set; } // 課規校部定
        public string GPRequired { get; set; } // 課規必選修
        public string GPCredit { get; set; } // 課規學分
        public string GPSYSubjectName { get; set; } // 課規指定學年度科目名稱

        public string SubjectLevelNew { get; set; } // 新科目級別

        public bool IsSubjectLevelChanged = false; // 科目級別是否有變更
    }
}
