using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHGraduationWarning.DAO
{
    // 處理指定學年科目名稱
    public class StudSchoolYearSubjectNameInfo
    {
        // 課程規劃表編號
        public string CoursePlanID { get; set; }
        // 課程規劃名稱
        public string CoursePlanName { get; set; }
        // 科目名稱
        public string SubjectName { get; set; }
        // 領域
        public string Domain { get; set; }
        // 分項
        public string Entry { get; set; }
        // 校部定
        public string RequiredBy { get; set; }
        // 必選修
        public string Required { get; set; }
        // 學分數
        public string Credit { get; set; }
        // 指定學年科目名稱
        public string SchoolYearSubjectName { get; set; }
        // 課程規劃指定學年科目名稱
        public string CoursePlanSchoolYearSubjectName { get; set; }
        // 問題說明
        public string ProblemDescription { get; set; }
    }
}
