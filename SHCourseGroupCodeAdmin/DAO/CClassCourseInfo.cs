using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace SHCourseGroupCodeAdmin.DAO
{
    /// <summary>
    /// 一般對開開課使用
    /// </summary>
    public class CClassCourseInfo
    {
        public string ClassID { get; set; }

        public string ClassName { get; set; }

        public string GradeYear { get; set; }
        // 班上學生，加入修課學生用
        public List<string> RefStudentIDList = new List<string>();

        /// <summary>
        /// 原始XML
        /// </summary>
        public XElement RefGPlanXML;

        /// <summary>
        /// 原班
        /// </summary>
        public List<XElement> OpenSubjectSourceList = new List<XElement>();

        /// <summary>
        /// 原班對開
        /// </summary>
        public List<XElement> OpenSubjectSourceBList = new List<XElement>();

        // 對開課程設定
        public Dictionary<string, bool> SubjectBDict = new Dictionary<string, bool>();

        /// <summary>
        /// 新增課程名稱
        /// </summary>
        public List<string> AddCourseNameList = new List<string>();

        /// <summary>
        /// 參考課程規劃ID
        /// </summary>
        public string RefGPID { get; set; }

        public string RefGPName { get; set; }
    }
}
