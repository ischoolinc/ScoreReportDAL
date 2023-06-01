using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
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

        // 科目名稱+科目級別,科目資料(存放可補修+可重修 科目名稱+級別過濾後清單)
        public Dictionary<string, XmlElement> dicRetake = new Dictionary<string, XmlElement>();
        // 規則,科目名稱+級別(存放符合可補修+可重修,規則與科目+級別清單)
        public Dictionary<string, List<string>> dicRetaleRelate = new Dictionary<string, List<string>>();

        // 畢業審查是否通過
        public string GraduationCheck = "";

        // 核心科目規則對照
        public Dictionary<string, string> dicCoreSubjectRule = new Dictionary<string, string>();

    }
}
