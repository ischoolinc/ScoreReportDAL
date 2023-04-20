using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHSemsSubjectCheckEdit.DAO
{
    public class GPlanSubjectInfo
    {
        //  科目名稱
        public string SubjectName { get; set; }

        //  科目級別
        public string SubjectLevel { get; set; }

        // 年級
        public string GradeYear { get; set; }

        // 學期
        public string Semester { get; set; }

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
        public string SubjectNameByYear { get; set; }

        // 課程代碼
        public string CourseCode { get; set; }
    }
}
