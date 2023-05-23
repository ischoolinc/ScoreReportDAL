using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace SHGraduationWarning.DAO
{
    // 畢業預警報表使用
    public class ReportStudentInfo
    {
        // 學生系統編號
        public string StudentID { get; set; }
        // 學號
        public string StudentNumber { get; set; }
        // 班級
        public string ClassName { get; set; }
        // 座號
        public string SeatNo { get; set; }
        // 姓名
        public string StudentName { get; set; }
        // 科別
        public string DeptName { get; set; }

        // 班級系統編號
        public string ClassID { get; set; }

        // 科別系統編號
        public string DeptID { get; set; }

        // 畢業檢查資料
        public XmlElement GraGrandCheckXml = null;

    }
}
