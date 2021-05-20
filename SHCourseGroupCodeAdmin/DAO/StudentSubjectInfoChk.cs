using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseGroupCodeAdmin.DAO
{
    /// <summary>
    /// 學生學期科目成績檢查
    /// </summary>
   public  class StudentSubjectInfoChk
    {
        /// <summary>
        /// 學生系統編號
        /// </summary>
        public string StudentID { get; set; }

        /// <summary>
        /// 學號
        /// </summary>
        public string StudentNumber { get; set; }

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
        public string StudentName { get; set; }

        /// <summary>
        /// GDC code
        /// </summary>
        public string gdc_code { get; set; }

        /// <summary>
        /// 檢查用學期科目成績
        /// </summary>
        public List<SubjectInfoChk> SubjectInfoChkList = new List<SubjectInfoChk>();

    }
}
