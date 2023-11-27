using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHGraduationWarning.DAO
{
    public class CourseInfo
    {
        // 學年度 
        public string SchoolYear { get; set; }

        // 學期
        public string Semester { get; set; }
        
        // 課程系統編號
        public string CourseID { get; set; }

        // 課程名稱
        public string CourseName { get; set; }

        // 科目名稱
        public string SubjectName { get; set; }

        // 科目級別
        public string SubjectLevel { get; set; }

        // 分項類別
        public string Entry { get; set; }

        // 領域
        public string Domain { get; set; }

        // 必選修
        public string Required { get; set; }

        // 校部訂 
        public string RequiredBy { get; set; }

        // 學分
        public string Credit { get; set; }

        // 不計入學分
        public string NotIncludedInCredit { get; set; }

        // 不評分
        public string NotIncludedInCalc { get; set; }

        // 新科目名稱
        public string NewSubjectName { get; set; }


        // 新科目級別
        public string NewSubjectLevel { get; set; }

        // 使用課程規劃表名稱
        public string GraduationPlanName { get; set; }

        // 錯誤訊息
        public string ErrorMessage { get; set; }

    }
}
