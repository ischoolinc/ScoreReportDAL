using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseGroupCodeAdmin.DAO
{
    public class CourseInfoChk
    {
        /// <summary>
        /// 課程系統編號
        /// </summary>
        public string CourseID { get; set; }
        /// <summary>
        /// 學年度
        /// </summary>
        public string SchoolYear { get; set; }
        /// <summary>
        ///   學期
        /// </summary>
        public string Semester { get; set; }
        /// <summary>
        /// 課程名稱
        /// </summary>
        public string CourseName { get; set; }

        /// <summary>
        /// 科目名稱
        /// </summary>
        public string SubjectName { get; set; }
        /// <summary>
        /// 科目級別
        /// </summary>
        public string SubjectLevel { get; set; }
        /// <summary>
        /// 部定校訂
        /// </summary>
        public string RequireBy { get; set; }
        /// <summary>
        /// 必修選修
        /// </summary>
        public string IsRequired { get; set; }  
        /// <summary>
        /// 學分數
        /// </summary>
        public string Credit { get; set; }   
        /// <summary>
        /// 節數
        /// </summary>
        public string Period { get; set; }
        /// <summary>
        /// 課程代碼
        /// </summary>
        public string course_code { get; set; }
        /// <summary>
        /// 授課學期學分節數
        /// </summary>
        public string credit_period { get; set; }
        /// <summary>
        /// Memo
        /// </summary>
        public string Memo { get; set; }

        /// <summary>
        /// 群科班代碼
        /// </summary>
        public string GroupCode { get; set; }
    }
}
