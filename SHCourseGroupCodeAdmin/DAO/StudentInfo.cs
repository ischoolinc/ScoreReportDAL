using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseGroupCodeAdmin
{
    /// <summary>
    /// 產出學生資訊
    /// </summary>
    public class StudentInfo
    {
        /// <summary>
        /// 學生系統編號
        /// </summary>
        public string StudentID { get; set; }

        /// <summary>
        /// 身分證號
        /// </summary>
        public string IDNumber { get; set; }

        /// <summary>
        /// 班級
        /// </summary>
        public string ClassName { get; set; }

        /// <summary>
        /// 座號
        /// </summary>
        public string SeatNo { get; set; }

        /// <summary>
        /// 姓名
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 科別
        /// </summary>
        public string Dept { get; set; }

        /// <summary>
        /// 學業(原始)成績
        /// </summary>
        public string EntryScore { get; set; }
    }
}
