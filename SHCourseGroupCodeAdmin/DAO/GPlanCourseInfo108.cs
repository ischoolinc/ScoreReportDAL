using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseGroupCodeAdmin.DAO
{
    public class GPlanCourseInfo108
    {
        /// <summary>
        /// 課程名稱
        /// </summary>
        public string CourseName { get; set; }
        /// <summary>
        /// 科目名稱
        /// </summary>
        public string SubjectName { get; set; }
        /// <summary>
        /// 必選修
        /// </summary>
        public string Required { get; set; }
        /// <summary>
        /// 部定校訂
        /// </summary>
        public string RequiredBy { get; set; }
        /// <summary>
        /// 學分
        /// </summary>
        public string Credit { get; set; }
        /// <summary>
        /// 不需計算
        /// </summary>
        public bool NotIncludedInCalc { get; set; }
        /// <summary>
        /// 不需學分 (False 表示需要算學分)
        /// </summary>
        public bool NotIncludedInCredit { get; set; }
        /// <summary>
        /// 分項
        /// </summary>
        public string Entry { get; set; }
        /// <summary>
        /// 群組代碼
        /// </summary>
        public string GroupCode { get; set; }
        /// <summary>
        /// 群組名稱
        /// </summary>
        public string GroupName { get; set; }
        /// <summary>
        /// 課程代碼
        /// </summary>
        public string CourseCode { get; set; }

        /// <summary>
        /// 授課學期學分_節數
        /// </summary>
        public string credit_period { get; set; }

        /// <summary>
        /// 說明
        /// </summary>
        public string Memo { get; set; }

        public string GradeYear { get; set; }

    }
}
