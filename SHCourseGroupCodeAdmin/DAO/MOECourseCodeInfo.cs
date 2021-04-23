using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseGroupCodeAdmin.DAO
{
    /// <summary>
    /// 課程代碼大表資料
    /// </summary>
    public class MOECourseCodeInfo
    {       
        /// <summary>
        /// 群組代碼
        /// </summary>
        public string group_code { get; set; }

        /// <summary>
        /// 課程代碼
        /// </summary>
        public string course_code { get; set; }

        /// <summary>
        /// 科目名稱
        /// </summary>
        public string subject_name { get; set; }

        /// <summary>
        /// 入學年      
        /// </summary>
        public string entry_year { get; set; }
        /// <summary>
        /// 部定校訂
        /// </summary>
        public string require_by { get; set; }
        /// <summary>
        /// 必修選修
        /// </summary>
        public string is_required { get; set; }
        /// <summary>
        /// 課程類型
        /// </summary>
        public string course_type { get; set; }
        /// <summary>
        /// 群別
        /// </summary>
        public string group_type { get; set; }

        /// <summary>
        /// 科別
        /// </summary>
        public string subject_type { get; set; }
        /// <summary>
        /// 班群
        /// </summary>
        public string class_type { get; set; }
        /// <summary>
        /// 授課學期學分_節數
        /// </summary>
        public string credit_period { get; set; }
    }
}
